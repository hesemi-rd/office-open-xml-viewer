import { computeVisibleRange, PT_TO_PX, type VisibleRange } from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';
import type { RenderPageOptions } from './types';

/**
 * Options for {@link DocxScrollViewer}. Extends `RenderPageOptions` (per-page
 * render knobs) and `LoadOptions` (parse/worker knobs). See design §8.1.
 */
export interface DocxScrollViewerOptions extends RenderPageOptions, LoadOptions {
  /** Base fit width in CSS px → base zoom scale. Default: the container's width
   *  at first non-zero layout (design §7/§11 zero-width deferral). */
  width?: number;
  /** Vertical gap (px) between consecutive pages. Default 16. */
  gap?: number;
  /** Pages kept mounted beyond the viewport on each side. Default 1. */
  overscan?: number;
  /** Per-page transparent text-selection overlay. MAIN render mode only:
   *  in worker mode `onTextRun` cannot cross the worker boundary, so the overlay
   *  stays empty and the viewer logs one warning (design §11). */
  enableTextSelection?: boolean;
  /** Minimum zoom scale (px-per-pt multiplier floor). Default 0.1. */
  zoomMin?: number;
  /** Maximum zoom scale. Default 4. */
  zoomMax?: number;
  /** Enable `Ctrl`/`Cmd`+wheel zoom. Default true. */
  enableZoom?: boolean;
  /**
   * Inject an already-loaded engine to share one parse across panes (design §14).
   * When set: `load()` is unsupported (throws), the engine's own `mode` wins (an
   * explicitly conflicting `opts.mode` throws at construction, design §11), and
   * `destroy()` does NOT destroy this engine (the caller owns its lifecycle).
   */
  document?: DocxDocument;
  /** Fires when the top-most visible page changes. `topIndex` from
   *  `computeVisibleRange` (the first page intersecting the viewport top,
   *  EXCLUDING overscan). */
  onVisiblePageChange?: (topIndex: number, total: number) => void;
  /** Error callback. When set, `load()` invokes it and resolves; otherwise the
   *  error is rethrown (shared viewer error contract). */
  onError?: (err: Error) => void;
}

/** One mounted page. `canvas` is the drawn page; `textLayer` the optional
 *  per-page selection overlay (main mode only). `renderedPage` guards against
 *  re-rendering a recycled slot for a page whose render is still in flight. */
interface PageSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  /** page index this slot is currently rendering / has rendered, or -1 when free. */
  renderedPage: number;
  /** worker-mode: the last ImageBitmap painted into this slot, to `.close()` on recycle. */
  bitmap: ImageBitmap | null;
  /** bitmaprenderer ctx (worker mode), grabbed once per canvas. */
  bitmapCtx: ImageBitmapRenderingContext | null;
}

export class DocxScrollViewer {
  private _doc: DocxDocument | null = null;
  private readonly _injected: boolean;
  private readonly _opts: DocxScrollViewerOptions;
  private readonly _container: HTMLElement;
  private readonly _wrapper: HTMLDivElement;
  private readonly _scrollHost: HTMLDivElement;
  private readonly _spacer: HTMLDivElement;
  /** Resolved render mode. When an engine is injected the engine's own `mode`
   *  is authoritative (design §11 — no silent mis-pathing / no probing); an
   *  explicitly conflicting `opts.mode` is rejected at construction. When self-
   *  loading, `opts.mode` decides and `load()` passes it to `DocxDocument.load`. */
  private _mode: 'main' | 'worker';

  /** px-per-pt zoom multiplier. Base fit maps the first page's width to the
   *  container width (or opts.width). Zoom multiplies this (design §7). */
  private _scale = 1;
  /** Whether the base fit scale has been established. Set true the first time
   *  `relayout()` resolves a positive base scale. We use an explicit flag rather
   *  than a `_scale === 1` sentinel because a fit scale of exactly 1 is a valid
   *  established state (a 1× fit would otherwise be re-fit forever). */
  private _scaleEstablished = false;
  /** Live slots keyed by page index. */
  private readonly _slots = new Map<number, PageSlot>();
  /** Recyclable detached slots (canvas + textLayer reused across pages). */
  private readonly _free: PageSlot[] = [];
  /** Cached per-page heights in px at the current scale (index-aligned). */
  private _heights: number[] = [];
  private _lastRange: VisibleRange | null = null;
  private _lastTopIndex = -1;
  private _scrollListener: (() => void) | null = null;
  /** Worker mode: page indices whose bitmap render is currently dispatched to the
   *  engine. Coalesces a scroll storm — we never dispatch a second render for a
   *  page whose first is still in flight — and lets us drop pages that scrolled
   *  out of the window before dispatch (design §11 worker coalescing). */
  private readonly _bitmapInFlight = new Set<number>();

  constructor(container: HTMLElement, opts: DocxScrollViewerOptions = {}) {
    this._container = container;
    this._opts = opts;
    this._injected = !!opts.document;
    if (this._injected) {
      const engine = opts.document as DocxDocument;
      // Injected engine ⇒ its own mode is the fact (design §11). An EXPLICITLY
      // conflicting opts.mode is a mis-configuration and is rejected here; an
      // absent opts.mode is fine.
      if (opts.mode !== undefined && opts.mode !== engine.mode) {
        throw new Error(
          `DocxScrollViewer: opts.mode='${opts.mode}' conflicts with the injected engine's mode='${engine.mode}'. ` +
            'Omit opts.mode when injecting an engine — the engine owns its render mode.',
        );
      }
      this._doc = engine;
      this._mode = engine.mode;
    } else {
      this._mode = opts.mode ?? 'main';
    }

    // container → wrapper → scrollHost → spacer  (design §6)
    this._wrapper = document.createElement('div');
    this._wrapper.style.cssText = 'position:relative;width:100%;height:100%;overflow:hidden;';
    this._scrollHost = document.createElement('div');
    this._scrollHost.style.cssText = 'position:absolute;inset:0;overflow:auto;';
    this._spacer = document.createElement('div');
    this._spacer.style.cssText = 'position:absolute;top:0;left:0;width:1px;height:0;pointer-events:none;';
    this._scrollHost.appendChild(this._spacer);
    this._wrapper.appendChild(this._scrollHost);
    this._container.appendChild(this._wrapper);

    this._scrollListener = () => this._onScroll();
    this._scrollHost.addEventListener('scroll', this._scrollListener);

    if (this._injected) {
      // An injected engine is already loaded, so lay out + mount the first
      // window immediately. relayout() is idempotent and defers under a
      // zero-width container (the resize path — T6 — re-runs it once width
      // appears).
      this.relayout();
    }
  }

  /**
   * Load a DOCX from URL or ArrayBuffer and render the first window.
   * UNSUPPORTED when an engine was injected via `opts.document` (throws) — the
   * caller already owns the parsed engine.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    if (this._injected) {
      throw new Error(
        'DocxScrollViewer.load() is unsupported when an engine is injected via opts.document; the injected engine is already loaded.',
      );
    }
    try {
      this._doc = await DocxDocument.load(source, {
        useGoogleFonts: this._opts.useGoogleFonts,
        maxZipEntryBytes: this._opts.maxZipEntryBytes,
        math: this._opts.math,
        mode: this._mode,
      });
      // Layout + first render wired in T2/T3.
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      if (this._opts.onError) {
        this._opts.onError(e);
        return;
      }
      throw e;
    }
  }

  get pageCount(): number {
    return this._doc?.pageCount ?? 0;
  }

  /** CSS px width of page `i` at the current scale. */
  private _pageWidthPx(i: number): number {
    return this._doc!.pageSize(i).widthPt * PT_TO_PX * this._scale;
  }

  /** CSS px height of page `i` at the current scale. */
  private _pageHeightPx(i: number): number {
    return this._doc!.pageSize(i).heightPt * PT_TO_PX * this._scale;
  }

  /** The container (fit) width, deferring when the container is unlaid-out. */
  private _fitWidthPx(): number {
    if (this._opts.width && this._opts.width > 0) return this._opts.width;
    const cw = this._container.clientWidth || this._scrollHost.clientWidth;
    return cw > 0 ? cw : 0; // 0 ⇒ defer (design §11 zero-width deferral)
  }

  /** Base scale: first page's width fit to the fit-width. Returns 0 when the
   *  container has no width yet (deferral). */
  private _baseScale(): number {
    if (!this._doc || this._doc.pageCount === 0) return 0;
    const w = this._fitWidthPx();
    if (w <= 0) return 0;
    const firstWpt = this._doc.pageSize(0).widthPt;
    if (firstWpt <= 0) return 0;
    return w / (firstWpt * PT_TO_PX);
  }

  /** Recompute heights + spacer, then mount the visible window. Called after
   *  load, resize, and zoom. Idempotent. */
  relayout(): void {
    if (!this._doc) return;
    // Establish the base fit scale on the first layout that has a positive
    // width. Zoom (T4) layers its own multiplier on top of this; here we only
    // set the base. An explicit `_scaleEstablished` flag (NOT a `_scale === 1`
    // sentinel) so a legitimate 1× fit is not re-fit on every relayout.
    if (!this._scaleEstablished) {
      const base = this._baseScale();
      if (base > 0) {
        this._scale = base;
        this._scaleEstablished = true;
      } else {
        return; // container has no width yet — retry on the next resize
      }
    }
    this._recomputeHeights();
    this._syncSpacer();
    this._mountVisible();
  }

  private _recomputeHeights(): void {
    const n = this._doc!.pageCount;
    const h = new Array<number>(n);
    for (let i = 0; i < n; i++) h[i] = this._pageHeightPx(i);
    this._heights = h;
  }

  private _gap(): number {
    return this._opts.gap ?? 16;
  }

  private _overscan(): number {
    return this._opts.overscan ?? 1;
  }

  private _range(): VisibleRange {
    return computeVisibleRange(
      this._heights,
      this._gap(),
      this._scrollHost.scrollTop,
      this._scrollHost.clientHeight,
      this._overscan(),
    );
  }

  private _syncSpacer(): void {
    const r = this._range();
    this._lastRange = r;
    this._spacer.style.height = `${r.totalHeight}px`;
  }

  private _onScroll(): void {
    if (!this._doc || !this._scaleEstablished) return;
    this._mountVisible();
  }

  /** Mount/recycle slots for the current visible window. */
  private _mountVisible(): void {
    if (!this._doc || this._doc.pageCount === 0) return;
    const r = this._range();
    this._lastRange = r;

    // Detach slots that left [start, end] into the free pool.
    for (const [idx, slot] of [...this._slots]) {
      if (idx < r.start || idx > r.end) {
        this._recycleSlot(idx, slot);
      }
    }
    // Mount any missing index in the window.
    for (let i = r.start; i <= r.end; i++) {
      if (!this._slots.has(i)) {
        const slot = this._acquireSlot();
        this._positionSlot(slot, i, r);
        this._slots.set(i, slot);
        this._renderSlot(i, slot); // T3
      } else {
        // Re-position (offsets shift after a spacer/height change).
        this._positionSlot(this._slots.get(i)!, i, r);
      }
    }
    // onVisiblePageChange only on change (T6 refines; wired here for topIndex).
    if (r.topIndex !== this._lastTopIndex) {
      this._lastTopIndex = r.topIndex;
      this._opts.onVisiblePageChange?.(r.topIndex, this._doc.pageCount);
    }
  }

  private _acquireSlot(): PageSlot {
    const reused = this._free.pop();
    if (reused) {
      reused.renderedPage = -1;
      this._scrollHost.appendChild(reused.wrapper);
      return reused;
    }
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:absolute;left:0;right:0;margin:0 auto;';
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block;background:#fff;';
    wrapper.appendChild(canvas);
    let textLayer: HTMLDivElement | null = null;
    const bitmapCtx: ImageBitmapRenderingContext | null = null;
    if (this._opts.enableTextSelection) {
      textLayer = document.createElement('div');
      textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      wrapper.appendChild(textLayer);
    }
    this._scrollHost.appendChild(wrapper);
    const slot: PageSlot = { wrapper, canvas, textLayer, renderedPage: -1, bitmap: null, bitmapCtx };
    return slot;
  }

  private _recycleSlot(idx: number, slot: PageSlot): void {
    this._slots.delete(idx);
    // Close any worker bitmap held by this slot (T3 sets slot.bitmap).
    if (slot.bitmap) {
      slot.bitmap.close();
      slot.bitmap = null;
    }
    slot.renderedPage = -1;
    slot.wrapper.remove();
    this._free.push(slot);
  }

  private _positionSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    slot.wrapper.style.top = `${r.offsets[i]}px`;
    const wpx = this._pageWidthPx(i);
    const hpx = this._pageHeightPx(i);
    slot.wrapper.style.width = `${wpx}px`;
    slot.wrapper.style.height = `${hpx}px`;
  }

  /** Device-pixel ratio for a render (opts override → window → 1). */
  private _dpr(): number {
    return this._opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
  }

  /**
   * Render page `i` into `slot`. Routes strictly on the constructor-resolved
   * `_mode` (design §11 — no probing, no silent mis-pathing): `main` ⇒ paint the
   * slot's canvas directly via `renderPage`; `worker` ⇒ transfer an ImageBitmap
   * from `renderPageToBitmap`.
   *
   * Slot-identity guard: a slot recycled to a DIFFERENT page while a previous
   * render is in flight must not repaint the stale page. `slot.renderedPage`
   * tracks the page this slot is committed to; we stamp it up-front and bail on
   * resolution if it changed (the engine's own token guard is per-canvas; this is
   * the viewer's per-slot page-identity check).
   */
  private _renderSlot(i: number, slot: PageSlot): void {
    if (!this._doc) return;
    // Slot-identity guard: this slot is already rendering / has rendered page i.
    if (slot.renderedPage === i) return;
    slot.renderedPage = i;

    const dpr = this._dpr();
    const widthPx = this._pageWidthPx(i);

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr);
      return;
    }

    // Main mode: render straight onto the slot's canvas.
    const runs: DocxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const onTextRun = wantOverlay ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
    this._doc
      .renderPage(slot.canvas, i, {
        width: widthPx, // this page's own px width → uniform px-per-pt scale (§7)
        dpr,
        defaultTextColor: this._opts.defaultTextColor,
        showTrackChanges: this._opts.showTrackChanges,
        onTextRun,
      })
      .then(() => {
        // A recycle may have re-purposed this slot for a different page (or freed
        // it) mid-render; bail without painting the stale page.
        if (this._slots.get(i) !== slot || slot.renderedPage !== i) return;
        if (wantOverlay && slot.textLayer) {
          buildDocxTextLayer(
            slot.textLayer,
            runs,
            slot.canvas.style.width || `${slot.canvas.width}px`,
            slot.canvas.style.height || `${slot.canvas.height}px`,
          );
        }
      })
      .catch((err: unknown) => {
        this._opts.onError?.(err instanceof Error ? err : new Error(String(err)));
      });
  }

  /**
   * Worker-mode slot render: dispatch `renderPageToBitmap`, transfer the result
   * via a per-slot `bitmaprenderer` context, and manage the ImageBitmap lifecycle.
   *
   * Coalescing / drop-stale (design §11):
   *  - Skip if page `i` is already in flight (a scroll storm won't double-dispatch).
   *  - Skip if page `i` already left the mounted window before dispatch.
   *  - On resolution, if `slot` is no longer THIS page's live slot (it recycled to
   *    another page, or page `i` re-mounted onto a DIFFERENT slot while this render
   *    was in flight), close the orphan bitmap and skip the paint. In that
   *    re-mount case a live slot for `i` still awaits a render, so once we clear
   *    the in-flight guard we re-dispatch it — a page that recycled and re-mounted
   *    mid-flight must never stay blank.
   */
  private async _renderSlotBitmap(i: number, slot: PageSlot, widthPx: number, dpr: number): Promise<void> {
    if (this._bitmapInFlight.has(i)) return; // coalesce: already dispatched
    // Drop-stale before dispatch: if this page already scrolled out of the
    // mounted window, don't dispatch at all.
    if (this._slots.get(i) !== slot) return;
    this._bitmapInFlight.add(i);
    // Grab the bitmaprenderer ctx ONCE per canvas — a canvas holds one context
    // type for its lifetime. A recycled canvas keeps the ctx grabbed on its
    // first worker render (bitmapCtx survives recycle), so we never re-getContext
    // a canvas that already has one (which would throw a type-flip in the DOM).
    if (!slot.bitmapCtx) {
      slot.bitmapCtx = slot.canvas.getContext('bitmaprenderer') as ImageBitmapRenderingContext | null;
    }
    try {
      const bmp = await this._doc!.renderPageToBitmap(i, {
        width: widthPx,
        dpr,
        showTrackChanges: this._opts.showTrackChanges,
      });
      // The slot may have recycled to a different page while the worker ran, or
      // page `i` may have re-mounted onto a DIFFERENT slot. Either way this bitmap
      // is stale for THIS slot: close it and skip the paint (drop-stale, §11).
      if (this._slots.get(i) !== slot || slot.renderedPage !== i) {
        bmp.close();
        return;
      }
      // Close the slot's previous bitmap, then hold the new one BEFORE transfer so
      // a concurrent recycle between now and transfer can close it (the recycle
      // path closes slot.bitmap). transferFromImageBitmap consumes the bitmap, so
      // we null the field immediately after — leaving nothing to double-close.
      if (slot.bitmap) slot.bitmap.close();
      slot.bitmap = bmp;
      slot.canvas.width = bmp.width;
      slot.canvas.height = bmp.height;
      slot.canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
      slot.canvas.style.height = `${Math.round(bmp.height / dpr)}px`;
      slot.bitmapCtx?.transferFromImageBitmap(bmp);
      slot.bitmap = null; // transfer consumed it
    } catch (err) {
      this._opts.onError?.(err instanceof Error ? err : new Error(String(err)));
    } finally {
      this._bitmapInFlight.delete(i);
      // If a LIVE slot for page `i` still awaits its render (page `i` re-mounted
      // onto a different slot while this render was in flight, so the re-mount's
      // dispatch was coalesced away), issue it now that the guard is clear.
      const live = this._slots.get(i);
      if (live && live !== slot && !this._bitmapInFlight.has(i)) {
        // live.renderedPage === i already (set by _renderSlot on mount); the
        // dispatch was swallowed by the in-flight guard, not the identity guard.
        void this._renderSlotBitmap(i, live, this._pageWidthPx(i), this._dpr());
      }
    }
  }

  get topVisiblePage(): number {
    return this._lastRange?.topIndex ?? 0;
  }

  /** @internal test hook: page indices currently mounted. */
  mountedPageIndicesForTest(): number[] {
    return [...this._slots.keys()];
  }

  /**
   * Tear down the viewer: remove the DOM subtree and (only for a self-loaded
   * engine) destroy the engine. An injected engine is left intact — the caller
   * owns its lifecycle. Per-slot worker ImageBitmaps are closed on recycle.
   */
  destroy(): void {
    if (this._scrollListener) {
      this._scrollHost.removeEventListener('scroll', this._scrollListener);
      this._scrollListener = null;
    }
    for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
    this._free.length = 0;
    if (!this._injected) {
      this._doc?.destroy();
    }
    this._doc = null;
    this._wrapper.remove();
  }
}
