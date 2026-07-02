import { computeVisibleRange, PT_TO_PX, zoomStepScale, type VisibleRange } from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';
import type { RenderPageOptions } from './types';

/**
 * Options for {@link DocxScrollViewer}. Extends `RenderPageOptions` (per-page
 * render knobs, minus `onTextRun`) and `LoadOptions` (parse/worker knobs). See
 * design §8.1.
 *
 * `onTextRun` is omitted deliberately: the viewer drives it internally per
 * mounted slot to build the optional per-page selection overlay (gated by
 * `enableTextSelection`), so exposing it here would let a caller's callback be
 * silently overridden.
 */
export interface DocxScrollViewerOptions extends Omit<RenderPageOptions, 'onTextRun'>, LoadOptions {
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
  /** Error callback. When set, `load()` invokes it and resolves (otherwise the
   *  error is rethrown — shared viewer error contract). It ALSO fires for async
   *  per-slot render failures (both main `renderPage` and worker
   *  `renderPageToBitmap` rejections); a failed page is left blank rather than
   *  crashing the loop. Without an `onError`, render failures are logged via
   *  `console.error` so they are never fully silent. */
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
  /** worker-mode: a transient hold on a just-received ImageBitmap, set only
   *  between receipt from the worker and its `transferFromImageBitmap` (which
   *  consumes it, after which we null the field). Its purpose is the throw path:
   *  if the transfer throws, `destroy()`/`_recycleSlot` can still find and
   *  `.close()` the bitmap. Normally null once transfer completes. */
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
  /** Set by `destroy()`. Async render callbacks (main + worker) check it before
   *  reporting an error so a rejection that lands after teardown is swallowed
   *  rather than surfaced to a `onError` on a dead viewer. */
  private _destroyed = false;
  /** Worker mode: page indices whose bitmap render is currently dispatched to the
   *  engine. Coalesces a scroll storm — we never dispatch a second render for a
   *  page whose first is still in flight — and lets us drop pages that scrolled
   *  out of the window before dispatch (design §11 worker coalescing).
   *
   *  T4 ZOOM HAZARD (RESOLVED by the render epoch below): coalescing keys on page
   *  INDEX only, with no notion of the scale a dispatch was made at. Once
   *  `setScale` can change the zoom mid-flight, an in-flight bitmap dispatched at
   *  the OLD scale can still pass the on-resolution identity check if the SAME
   *  slot object is re-mounted for page `i` (the pool reuses slot objects, so
   *  `_slots.get(i) === slot && slot.renderedPage === i` can hold for an old
   *  dispatch), and get painted at the WRONG resolution. We fix this with a render
   *  epoch (`_renderEpoch`): each dispatch captures the epoch, and on resolution a
   *  moved epoch ⇒ STALE (close + re-dispatch the live slot). See
   *  `_renderSlotBitmap`. */
  private readonly _bitmapInFlight = new Set<number>();
  /** Render generation, bumped on every effective `setScale` (and the T6 resize
   *  re-fit, which routes through `setScale`). Stamped into each async render
   *  dispatch; a resolution whose captured epoch ≠ this value is STALE — its
   *  pixels/geometry are at a superseded scale. Worker path: close the orphan
   *  bitmap + re-dispatch the live slot. Main path: skip the (stale) text-layer
   *  build; the engine's per-canvas token already discards the stale pixels. */
  private _renderEpoch = 0;
  private _wheelListener: ((e: WheelEvent) => void) | null = null;

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

    // Ctrl/Cmd+wheel zoom (design §7). Bare wheel is left untouched so the
    // scrollHost scrolls natively. `enableZoom:false` installs no handler at all.
    // `{ passive: false }` is required because we call preventDefault() to stop
    // the browser's own ctrl+wheel page zoom.
    if (this._opts.enableZoom !== false) {
      this._wheelListener = (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) return; // bare wheel scrolls natively
        e.preventDefault();
        if (e.deltaY === 0) return;
        this.setScale(zoomStepScale(this._scale, e.deltaY));
      };
      this._scrollHost.addEventListener('wheel', this._wheelListener as EventListener, {
        passive: false,
      });
    }

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
      // Lay out + mount the first window now that the engine exists (mirrors the
      // injected-engine path in the constructor). relayout() is idempotent and
      // defers under a zero-width container — the resize path (T6) re-runs it
      // once width appears.
      this.relayout();
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
        this._renderSlot(i, slot);
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
      // _recycleSlot already reset renderedPage to -1 before pooling this slot.
      this._scrollHost.appendChild(reused.wrapper);
      return reused;
    }
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:absolute;left:0;right:0;margin:0 auto;';
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block;background:#fff;';
    wrapper.appendChild(canvas);
    let textLayer: HTMLDivElement | null = null;
    if (this._opts.enableTextSelection) {
      textLayer = document.createElement('div');
      textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      wrapper.appendChild(textLayer);
    }
    this._scrollHost.appendChild(wrapper);
    const slot: PageSlot = { wrapper, canvas, textLayer, renderedPage: -1, bitmap: null, bitmapCtx: null };
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
   *
   * Render epoch (main path): pixel staleness after a mid-flight `setScale` is
   * already handled by the engine's per-canvas token (the newer renderPage on the
   * same canvas wins) — `setScale` recycles + re-mounts, and the re-mount always
   * re-dispatches `renderPage` (renderedPage reset to -1), so a fresh render is
   * always issued. But the viewer-side side effects of a STALE resolution — the
   * text-layer build (its run geometry is at the OLD scale) and the renderedPage
   * bookkeeping — must NOT run, or a superseded render would rebuild the overlay
   * with stale x/y/w/h (the pool reuses slot objects, so the identity check alone
   * can pass for an old-epoch resolution). We gate them on the captured epoch.
   */
  private _renderSlot(i: number, slot: PageSlot): void {
    if (!this._doc) return;
    // Slot-identity guard: this slot is already rendering / has rendered page i.
    if (slot.renderedPage === i) return;
    slot.renderedPage = i;

    const dpr = this._dpr();
    const widthPx = this._pageWidthPx(i);
    const epoch = this._renderEpoch;

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
        // Stale if the epoch moved (a setScale rescaled mid-flight — the run
        // geometry is at the old scale), or a recycle re-purposed this slot for a
        // different page / freed it. Either way: skip the (stale) overlay build.
        // The engine's per-canvas token already discards the superseded pixels.
        if (epoch !== this._renderEpoch || this._slots.get(i) !== slot || slot.renderedPage !== i) return;
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
        this._reportRenderError(err);
      });
  }

  /** Route an async render failure to `onError`, or `console.error` when none is
   *  set (so failures are never fully silent), and never after teardown. */
  private _reportRenderError(err: unknown): void {
    if (this._destroyed) return;
    const e = err instanceof Error ? err : new Error(String(err));
    if (this._opts.onError) this._opts.onError(e);
    else console.error('[ooxml] DocxScrollViewer render failed:', e);
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
   *  - RENDER EPOCH: the dispatch captures `this._renderEpoch`. `setScale` bumps
   *    the epoch, so a resolution whose captured epoch ≠ the live epoch is STALE
   *    even when the SAME slot object is still mounted for page `i` (the pool
   *    reuses slot objects, so the identity check alone can't catch a zoom that
   *    happened mid-flight). A moved epoch ⇒ close the orphan + re-dispatch the
   *    live slot at the new scale, never paint the old-scale bitmap.
   */
  private async _renderSlotBitmap(i: number, slot: PageSlot, widthPx: number, dpr: number): Promise<void> {
    if (this._bitmapInFlight.has(i)) return; // coalesce: already dispatched
    // Drop-stale before dispatch: if this page already scrolled out of the
    // mounted window, don't dispatch at all.
    if (this._slots.get(i) !== slot) return;
    const epoch = this._renderEpoch;
    this._bitmapInFlight.add(i);
    // Whether this invocation actually painted its slot. When it did NOT (stale
    // epoch or moved identity), the `finally` may need to re-dispatch a live slot.
    let painted = false;
    // Grab the bitmaprenderer ctx ONCE per canvas — a canvas holds one context
    // type for its lifetime. A recycled canvas keeps the ctx grabbed on its
    // first worker render (bitmapCtx survives recycle), so we never re-getContext
    // a canvas that already has one. (getContext for a conflicting type returns
    // null rather than throwing; caching the first non-null ctx avoids relying on
    // that and skips redundant lookups.)
    if (!slot.bitmapCtx) {
      slot.bitmapCtx = slot.canvas.getContext('bitmaprenderer') as ImageBitmapRenderingContext | null;
    }
    try {
      const bmp = await this._doc!.renderPageToBitmap(i, {
        width: widthPx,
        dpr,
        defaultTextColor: this._opts.defaultTextColor,
        showTrackChanges: this._opts.showTrackChanges,
      });
      // Stale if EITHER (a) the epoch moved (a setScale rescaled mid-flight, so
      // this bitmap is at a superseded resolution — this catches the case where
      // the SAME slot object is re-mounted for page `i`, which the identity check
      // below cannot), or (b) the slot recycled to a different page / page `i`
      // re-mounted onto a DIFFERENT slot. Either way: close + skip the paint.
      if (epoch !== this._renderEpoch || this._slots.get(i) !== slot || slot.renderedPage !== i) {
        bmp.close();
        return;
      }
      // Close any prior bitmap, then hold the new one on the slot BEFORE the
      // transfer. JS is single-threaded so nothing recycles between here and the
      // transfer; the hold's real value is the throw path — if
      // transferFromImageBitmap throws, `destroy()`/`_recycleSlot` can still find
      // and close this bitmap. transferFromImageBitmap consumes the bitmap, so we
      // null the field immediately after — leaving nothing to double-close.
      if (slot.bitmap) slot.bitmap.close();
      slot.bitmap = bmp;
      slot.canvas.width = bmp.width;
      slot.canvas.height = bmp.height;
      slot.canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
      slot.canvas.style.height = `${Math.round(bmp.height / dpr)}px`;
      slot.bitmapCtx?.transferFromImageBitmap(bmp);
      slot.bitmap = null; // transfer consumed it
      painted = true;
    } catch (err) {
      this._reportRenderError(err);
    } finally {
      this._bitmapInFlight.delete(i);
      // Re-dispatch ONLY when this invocation went stale — a LIVE slot for page
      // `i` still awaits a correct render and the reason we didn't paint was
      // staleness, not a render failure. The two staleness cases:
      //  - IDENTITY MOVED (`live !== slot`): page `i` re-mounted onto a DIFFERENT
      //    slot while we ran (the re-mount's own dispatch was coalesced away by
      //    the in-flight guard), so the live slot has no render in flight.
      //  - EPOCH MOVED (`epoch !== this._renderEpoch`): a `setScale` bumped the
      //    epoch mid-flight, so this bitmap was at a superseded scale. The live
      //    slot may be the SAME object reused from the pool, which the identity
      //    test alone would miss — the epoch test catches the same-slot case.
      // NO RETRY ON PLAIN REJECTION: when the slot is still live at the same epoch
      // and we simply failed (`renderPageToBitmap` rejected or the transfer threw),
      // `!painted` holds but BOTH staleness tests are false, so we do NOT
      // re-dispatch. Retrying a plain failure would loop unbounded (reject →
      // re-dispatch → reject → …); the onError contract is that "a failed page is
      // left blank" (see DocxScrollViewerOptions.onError), so we leave it blank.
      // Bounded epoch-then-reject: an epoch-moved re-dispatch captures the NEW
      // epoch, so if that fresh render then rejects at the still-current epoch,
      // both tests are false and it stops — no unbounded retry.
      const live = this._slots.get(i);
      if (
        !painted &&
        live &&
        (live !== slot || epoch !== this._renderEpoch) &&
        !this._bitmapInFlight.has(i) &&
        !this._destroyed
      ) {
        // live.renderedPage === i already (set by _renderSlot on mount); the fresh
        // dispatch runs at the CURRENT epoch/scale via _pageWidthPx(i).
        void this._renderSlotBitmap(i, live, this._pageWidthPx(i), this._dpr());
      }
    }
  }

  /**
   * Set the absolute px-per-pt zoom scale, clamped inline to
   * `[zoomMin ?? 0.1, zoomMax ?? 4]` (absolute bounds, XlsxViewer convention — NOT
   * multiples of the base fit; design §3 keeps the clamp in the viewer, not core),
   * then re-anchor VERTICALLY so the page currently under the viewport top stays
   * fixed. A no-op when nothing is loaded or when the clamped scale is unchanged.
   *
   * Re-anchor (written from scratch — XlsxViewer only re-anchors horizontally):
   * capture `top = topIndex` and the intra-page fraction `intraFrac` from the
   * CURRENT range BEFORE rescale; after recomputing heights at the new scale,
   * `newScrollTop = offsets'[top] + intraFrac × heights'[top]`, clamped to
   * `[0, totalHeight' − viewportHeight]`. Because a page's height scales linearly
   * with `_scale`, the same fractional position maps exactly to the new geometry.
   *
   * CAVEAT — base fit below the floor: `relayout()` sets `_scale = base` WITHOUT
   * clamping to `[zoomMin, zoomMax]`. If the base fit is below `zoomMin` (a wide
   * page in a narrow container), the initial scale sits under the floor, but once
   * the user zooms via `setScale` the clamp pins the minimum to `zoomMin`, so they
   * can no longer return below the floor to the original base fit through this API.
   */
  setScale(scale: number): void {
    if (!this._doc || this._doc.pageCount === 0 || !this._scaleEstablished) return;
    const zoomMin = this._opts.zoomMin ?? 0.1;
    const zoomMax = this._opts.zoomMax ?? 4;
    const next = Math.min(zoomMax, Math.max(zoomMin, scale));
    if (next === this._scale) return;

    // Capture the anchor from the CURRENT layout, before rescale.
    const r0 = this._range();
    const top = r0.topIndex;
    const h0 = this._heights[top] || 0;
    // intraFrac: the fraction of page `top` that has scrolled above the viewport
    // top. Clamp to [0,1] — a scrollTop inside the trailing gap after page `top`
    // is attributed to `top` by computeVisibleRange and would push intraFrac past
    // 1, which would drift the re-anchor into the gap; pin the page instead.
    let intraFrac = h0 > 0 ? (this._scrollHost.scrollTop - r0.offsets[top]) / h0 : 0;
    intraFrac = Math.min(1, Math.max(0, intraFrac));

    // Bump the render epoch BEFORE recycling/re-dispatching so any in-flight
    // render dispatched at the old scale is recognised as stale on resolution.
    this._renderEpoch++;

    // Rescale, recompute heights, resize the spacer to the new total height.
    this._scale = next;
    this._recomputeHeights();
    const r1 = computeVisibleRange(this._heights, this._gap(), 0, this._scrollHost.clientHeight, this._overscan());
    this._spacer.style.height = `${r1.totalHeight}px`;

    // Pin the same fractional position of the same page under the viewport top.
    const maxTop = Math.max(0, r1.totalHeight - this._scrollHost.clientHeight);
    const wantTop = (r1.offsets[top] ?? 0) + intraFrac * (this._heights[top] || 0);
    this._scrollHost.scrollTop = Math.min(maxTop, Math.max(0, wantTop));

    // Re-mount at the new geometry, forcing a re-render of every slot (its px
    // size changed under the new scale, so the cached canvas is stale).
    this._mountVisibleForceRerender();
  }

  /** Like `_mountVisible` but recycles every live slot first so each re-mounts and
   *  re-renders at the current scale. Used by `setScale` (and the T6 resize
   *  re-fit): a slot's canvas px size changes with the scale, so the previously
   *  drawn pixels are stale and every visible page must be redrawn. */
  private _mountVisibleForceRerender(): void {
    for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
    this._mountVisible();
  }

  get topVisiblePage(): number {
    return this._lastRange?.topIndex ?? 0;
  }

  /** @internal test hook: page indices currently mounted. */
  mountedPageIndicesForTest(): number[] {
    return [...this._slots.keys()];
  }

  /** @internal test hook: the current absolute px-per-pt scale. */
  scaleForTest(): number {
    return this._scale;
  }

  /** @internal test hook: the base fit scale (pre-zoom) at the current width. */
  baseScaleForTest(): number {
    return this._baseScale();
  }

  /**
   * Tear down the viewer: remove the DOM subtree and (only for a self-loaded
   * engine) destroy the engine. An injected engine is left intact — the caller
   * owns its lifecycle. Per-slot worker ImageBitmaps are closed on recycle.
   */
  destroy(): void {
    this._destroyed = true;
    if (this._scrollListener) {
      this._scrollHost.removeEventListener('scroll', this._scrollListener);
      this._scrollListener = null;
    }
    if (this._wheelListener) {
      this._scrollHost.removeEventListener('wheel', this._wheelListener as EventListener);
      this._wheelListener = null;
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
