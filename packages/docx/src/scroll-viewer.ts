import { computeVisibleRange, PT_TO_PX, zoomStepScale, type VisibleRange } from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';
import type { RenderPageOptions } from './types';

/**
 * Debounce window (ms) after the last `setScale` in a zoom burst before the
 * full-resolution settle re-render is dispatched (design §7 "Flicker-free zoom").
 *
 * This is a UI-INTERACTION-FEEL policy constant, NOT an ECMA-376 / ISO-29500
 * value: it exists only so a rapid wheel/pinch gesture (which fires dozens of
 * `setScale` calls) coalesces into a single high-res render at the end instead of
 * re-rendering per tick. Each `setScale` shows an immediate CSS preview (the
 * existing bitmap stretched) and resets this timer; the settle fires once the
 * gesture pauses for `ZOOM_SETTLE_MS`. Lower = snappier but more redundant renders
 * mid-gesture; higher = fewer renders but a longer soft-preview tail. Deliberately
 * duplicated per viewer (a one-line timing constant, not shared logic).
 */
const ZOOM_SETTLE_MS = 150;

/**
 * Default CSS `box-shadow` painted on every page canvas — the soft drop shadow a
 * PDF reader casts under each sheet (matches the Examples/recipe look, which the
 * scroll viewer now reproduces with zero config). See
 * {@link DocxScrollViewerOptions.pageShadow}.
 */
const DEFAULT_PAGE_SHADOW = '0 1px 3px rgba(0,0,0,0.2)';

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
  /** Desk padding (px) ABOVE the FIRST page — the margin a PDF reader leaves
   *  between the top of the scroll surface and the first sheet. Default: `gap`
   *  (uniform desk rhythm — the first page sits the same distance from the top as
   *  pages sit from each other). Pass `0` for a flush-top layout. */
  paddingTop?: number;
  /** Desk padding (px) BELOW the LAST page — the margin below the final sheet.
   *  Default: `gap`. Pass `0` for a flush-bottom layout. */
  paddingBottom?: number;
  /** Desk gutter (px) to the LEFT of the pages — the horizontal margin between the
   *  left edge of the scroll surface and a page sitting flush-left (i.e. once
   *  zoomed wide enough that centering no longer applies). Default: `gap` (uniform
   *  desk rhythm — the horizontal gutters match the vertical ones). It also shrinks
   *  the container-derived FIT width so a page sits inside the gutters at 100%
   *  (an EXPLICIT `opts.width` is the page's CSS-width contract and is NOT reduced;
   *  the gutters still apply around placement). Pass `0` for a flush-left layout. */
  paddingLeft?: number;
  /** Desk gutter (px) to the RIGHT of the pages. Default: `gap`. Shrinks the
   *  container-derived fit width symmetrically with `paddingLeft`. Pass `0` for a
   *  flush-right layout. */
  paddingRight?: number;
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
   * CSS `background` shorthand for the scroll surface (the "desk") visible
   * behind and between pages — the gray a PDF reader paints around the sheet.
   * Applied to the viewer-owned scroll host. The pages themselves are always
   * drawn on the document's own white canvas and are unaffected. Default
   * `undefined`: the scroll surface stays transparent so the host container's
   * background shows through (non-breaking).
   */
  background?: string;
  /**
   * CSS `box-shadow` painted on every page CANVAS (not the wrapper — the
   * text-selection overlay must not cast its own shadow). The soft drop shadow a
   * PDF reader leaves under each sheet.
   *
   * - Default (`undefined`): `'0 1px 3px rgba(0,0,0,0.2)'` — the recipe look, so
   *   the scroll viewer reproduces the Examples appearance with zero config.
   * - `false`: NO shadow (flat pages).
   * - A custom string is applied verbatim. A spread-only ring such as
   *   `'0 0 0 1px #c8ccd0'` gives a crisp 1px BORDER look — and because
   *   `box-shadow` never affects layout (unlike `border`, which would grow the
   *   box and shift every offset), a border and a drop shadow are the SAME knob
   *   here rather than two competing options.
   */
  pageShadow?: string | false;
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
  /** The `_scale` at which this slot's on-screen canvas bitmap (and text overlay)
   *  were last rendered, or -1 when unrendered. The flicker-free CSS preview
   *  (design §7) stretches that bitmap to the new layout size on `setScale` and
   *  scales the text overlay by `newScale / renderedScale`; the debounced settle
   *  re-render then repaints at the new scale and updates this to match. */
  renderedScale: number;
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
  /** Render generation, bumped on every effective `setScale` (and the resize
   *  re-fit in `_onResize`, which routes through `setScale`). Stamped into each async render
   *  dispatch; a resolution whose captured epoch ≠ this value is STALE — its
   *  pixels/geometry are at a superseded scale. Worker path: close the orphan
   *  bitmap + re-dispatch the live slot. Main path: skip the (stale) text-layer
   *  build; the engine's per-canvas token already discards the stale pixels. */
  private _renderEpoch = 0;
  /** Pending settle-render timer handle (design §7 mechanism 2). Set by
   *  `_scheduleSettle` after each `setScale`, reset on the next one so a burst
   *  dispatches ONE settle at the end, and cleared in `destroy()`. `ReturnType`
   *  of `setTimeout` (a number in the DOM, a Timeout object in node) so the type
   *  is host-agnostic. */
  private _settleTimer: ReturnType<typeof setTimeout> | null = null;
  private _wheelListener: ((e: WheelEvent) => void) | null = null;
  /** One-shot latch for the worker-mode text-selection warning. The overlay is a
   *  main-mode-only feature: in worker mode the per-run `onTextRun` geometry
   *  cannot cross the worker boundary, so an `enableTextSelection` overlay stays
   *  empty. We warn once (parity with `DocxViewer`) rather than per slot. */
  private _warnedNoTextSelection = false;
  /** Observes the container so a width change re-fits the base scale. Disconnected
   *  in `destroy()`. */
  private _resizeObserver: ResizeObserver | null = null;
  /** The base fit scale at the last established/re-fit layout. `_onResize` divides
   *  `_scale` by this to recover the current zoom multiplier so a width change
   *  re-fits the base while preserving the user's zoom (design §11). */
  private _prevBase = 0;
  /** The fit width (px) the base scale was last established at. Lets `_onResize`
   *  skip the re-fit when only the height changed (a ResizeObserver fires on ANY
   *  box change, but only a WIDTH change alters the fit-to-width base scale). */
  private _lastFitWidth = 0;
  /** Resolved page-canvas `box-shadow` (design: the recipe drop shadow by
   *  default). Resolved ONCE with `??` — NOT `||` — so `pageShadow: false`
   *  survives as the "no shadow" sentinel (a `||` would treat `false` as absent
   *  and wrongly re-apply the default). Applied by `_applyPageShadow` at EVERY
   *  canvas-creation site (`_acquireSlot` and the double-buffer spare in
   *  `_settleSlot`) so a recycled/re-mounted slot and a settle-swapped spare all
   *  carry it. */
  private readonly _pageShadow: string | false;

  constructor(container: HTMLElement, opts: DocxScrollViewerOptions = {}) {
    // A <canvas> is an HTMLElement too, so the type system cannot stop a caller
    // used to the pager API (DocxViewer takes a canvas) from passing one — but
    // canvas children never render, so the viewer would come up silently blank.
    // Fail loudly with the fix instead. (tagName, not instanceof: cross-realm safe.)
    if (container.tagName === 'CANVAS') {
      throw new Error(
        'DocxScrollViewer takes a container element (e.g. a <div>), not a <canvas> — ' +
          'the viewer creates and manages its own canvases. Pass a block container; ' +
          'for the single-page canvas API use DocxViewer.',
      );
    }
    this._container = container;
    this._opts = opts;
    // `??` (not `||`): a caller's explicit `false` must disable the shadow, not
    // fall through to the default.
    this._pageShadow = opts.pageShadow ?? DEFAULT_PAGE_SHADOW;
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
    // The "desk" behind/between pages. Undefined ⇒ transparent (container shows
    // through); pages keep their own white canvas regardless.
    if (opts.background) this._scrollHost.style.background = opts.background;
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

    // Re-fit the base scale on a container resize (design §11). A container that
    // is 0-wide at construction (a common flexbox/tab layout) establishes its
    // scale on the first non-zero resize — the zero-width deferral is completed
    // here. `ResizeObserver` may be absent in a non-DOM host; guard for it.
    if (typeof ResizeObserver !== 'undefined') {
      this._resizeObserver = new ResizeObserver(() => this._onResize());
      this._resizeObserver.observe(this._container);
    }

    if (this._injected) {
      // An injected engine is already loaded, so lay out + mount the first
      // window immediately. relayout() is idempotent and defers under a
      // zero-width container (the resize path re-runs it once width appears).
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
        workerTimeoutMs: this._opts.workerTimeoutMs,
        wasmUrl: this._opts.wasmUrl,
        math: this._opts.math,
        mode: this._mode,
      });
      // Lay out + mount the first window now that the engine exists (mirrors the
      // injected-engine path in the constructor). relayout() is idempotent and
      // defers under a zero-width container — `_onResize` re-runs it once width
      // appears.
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

  /** The fit width (px), deferring when the container is unlaid-out. An EXPLICIT
   *  `opts.width` is the page's CSS-width contract and is returned UNCHANGED (the
   *  gutters still apply around placement, not to the width). The container-derived
   *  default instead targets `containerWidth − padL − padR` so a page sits INSIDE
   *  the horizontal gutters at 100%. A non-positive result (gutters wider than the
   *  container) is treated as unlaid-out — the same deferral as a zero-width box. */
  private _fitWidthPx(): number {
    if (this._opts.width && this._opts.width > 0) return this._opts.width;
    const cw = this._container.clientWidth || this._scrollHost.clientWidth;
    if (cw <= 0) return 0; // 0 ⇒ defer (design §11 zero-width deferral)
    const { left, right } = this._padH();
    const fit = cw - left - right;
    return fit > 0 ? fit : 0; // gutters ≥ container ⇒ defer (same as zero-width)
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

  /**
   * Recompute per-page heights + the spacer and re-mount the visible window.
   *
   * The viewer already calls this automatically after `load()`, an injected
   * engine, a container resize, and a zoom, so most integrations never need it.
   * It is public as a deliberate escape hatch: if the host mutates the layout in
   * a way the `ResizeObserver` cannot observe (e.g. a CSS change on an ancestor
   * that resizes the container without a box-size event, or a font that finishes
   * loading after first paint), call `relayout()` to force a re-fit. Idempotent —
   * safe to call repeatedly, and a no-op while the container has zero width (the
   * fit is deferred until width appears, design §11).
   */
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
        this._prevBase = base;
        this._lastFitWidth = this._fitWidthPx();
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

  /** Desk padding fed to `computeVisibleRange`: `paddingTop`/`paddingBottom`,
   *  each defaulting to `gap` (uniform rhythm). Resolved here (not stored) to
   *  mirror `_gap()`/`_overscan()`, and consumed at EVERY `computeVisibleRange`
   *  call site so the padded offsets are the single source of geometry. */
  private _pad(): { leading: number; trailing: number } {
    const gap = this._gap();
    return { leading: this._opts.paddingTop ?? gap, trailing: this._opts.paddingBottom ?? gap };
  }

  /** Horizontal desk gutters: `paddingLeft`/`paddingRight`, each defaulting to
   *  `gap` (uniform rhythm — the horizontal gutters match the vertical padding).
   *  Consumed by `_fitWidthPx` (to shrink the container-derived fit), by
   *  `_positionSlot` (the flush-left floor), and by `_syncSpacer` (the spacer
   *  width). Resolved here (not stored) to mirror `_gap()`/`_pad()`. */
  private _padH(): { left: number; right: number } {
    const gap = this._gap();
    return { left: this._opts.paddingLeft ?? gap, right: this._opts.paddingRight ?? gap };
  }

  private _range(): VisibleRange {
    return computeVisibleRange(
      this._heights,
      this._gap(),
      this._scrollHost.scrollTop,
      this._scrollHost.clientHeight,
      this._overscan(),
      this._pad(),
    );
  }

  private _syncSpacer(): void {
    const r = this._range();
    this._lastRange = r;
    this._spacer.style.height = `${r.totalHeight}px`;
    this._syncSpacerWidth();
  }

  /** Horizontal scroll extent: the widest page (docx pages can differ in width)
   *  plus both gutters. A spacer NARROWER than the container never creates a
   *  scrollbar (scrollWidth = max(clientWidth, content)), so it is always safe to
   *  set — it only matters when a zoomed-in page grows past the viewport, where it
   *  gives the gutters something to scroll to on either side. Max over per-page
   *  widths so the extent covers the widest page in the document. Called from
   *  `_syncSpacer` and after every scale change (zoom / resize re-fit) so the
   *  extent tracks the current page px width. */
  private _syncSpacerWidth(): void {
    const { left, right } = this._padH();
    let maxW = 0;
    for (let i = 0; i < this._heights.length; i++) {
      const w = this._pageWidthPx(i);
      if (w > maxW) maxW = w;
    }
    this._spacer.style.width = `${maxW + left + right}px`;
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
    // onVisiblePageChange fires ONLY when the top visible page actually changes
    // (change-only latch; `_lastTopIndex` starts at -1 so the first layout fires
    // once for page 0). Every mount path — scroll, zoom, resize re-fit, and
    // scrollToPage — funnels through here, so navigation never double-fires.
    if (r.topIndex !== this._lastTopIndex) {
      this._lastTopIndex = r.topIndex;
      this._opts.onVisiblePageChange?.(r.topIndex, this._doc.pageCount);
    }
  }

  /** Apply the resolved page-canvas shadow (design: recipe drop shadow by
   *  default, `false` ⇒ none). Single source so `_acquireSlot` and the
   *  double-buffer spare in `_settleSlot` stay in lock-step — a spare that missed
   *  this would lose the shadow on the settle swap. `box-shadow` never affects
   *  layout, so this is safe to (re)set on a live/pooled canvas without shifting
   *  any offset. */
  private _applyPageShadow(canvas: HTMLCanvasElement): void {
    if (this._pageShadow !== false) canvas.style.boxShadow = this._pageShadow;
  }

  private _acquireSlot(): PageSlot {
    const reused = this._free.pop();
    if (reused) {
      // _recycleSlot already reset renderedPage to -1 before pooling this slot.
      this._scrollHost.appendChild(reused.wrapper);
      return reused;
    }
    // `left` is set explicitly per mount by `_positionSlot` (JS centering with a
    // left-gutter floor), so no CSS auto-centering (`left:0;right:0;margin:0 auto`)
    // here — it would fight the explicit `left`.
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:absolute;';
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block;background:#fff;';
    this._applyPageShadow(canvas);
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
    const slot: PageSlot = {
      wrapper,
      canvas,
      textLayer,
      renderedPage: -1,
      renderedScale: -1,
      bitmap: null,
      bitmapCtx: null,
    };
    return slot;
  }

  private _recycleSlot(idx: number, slot: PageSlot): void {
    this._slots.delete(idx);
    // Close any worker bitmap held by this slot (T3 sets slot.bitmap).
    if (slot.bitmap) {
      slot.bitmap.close();
      slot.bitmap = null;
    }
    // Clear the per-slot text overlay so a slot sitting in the free pool holds no
    // stale spans. buildDocxTextLayer also clears on its next build, but an
    // unrendered pooled slot never gets that build, and the detached spans would
    // otherwise linger; drop them here.
    if (slot.textLayer) {
      slot.textLayer.innerHTML = '';
      // Drop any preview transform so a pooled slot re-used for another page does
      // not inherit a stale scale() before its overlay is rebuilt.
      slot.textLayer.style.transform = '';
      slot.textLayer.style.transformOrigin = '';
    }
    slot.renderedPage = -1;
    slot.renderedScale = -1;
    slot.wrapper.remove();
    this._free.push(slot);
  }

  private _positionSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    slot.wrapper.style.top = `${r.offsets[i]}px`;
    const wpx = this._pageWidthPx(i);
    const hpx = this._pageHeightPx(i);
    slot.wrapper.style.width = `${wpx}px`;
    slot.wrapper.style.height = `${hpx}px`;
    // Horizontal placement (replaces the old CSS `left:0;right:0;margin:0 auto`
    // auto-centering, which cannot honour a left gutter). Centre the page in the
    // scroll viewport, but never let its left edge cross the left gutter: when the
    // page is narrower than the viewport it is centred (`(cw − pw)/2 > padL`); once
    // zoomed wider than the viewport the centre would go negative, so the floor
    // pins it at `padL` and the overflow scrolls right. Formula deliberately
    // duplicated per viewer (one line; not hoisted to core).
    const { left: padL } = this._padH();
    const cw = this._scrollHost.clientWidth;
    slot.wrapper.style.left = `${Math.max(padL, (cw - wpx) / 2)}px`;
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
    const scale = this._scale;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr, scale);
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
        // This fresh render defines the scale the on-screen bitmap now lives at,
        // so a subsequent zoom preview stretches from HERE.
        slot.renderedScale = scale;
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

  /** Warn once when an `enableTextSelection` overlay was requested but the render
   *  mode is `worker` (so the overlay stays empty). Same wording as
   *  `DocxViewer._render` — one warning per viewer, not per slot. */
  private _maybeWarnNoTextSelection(): void {
    if (this._opts.enableTextSelection && !this._warnedNoTextSelection) {
      this._warnedNoTextSelection = true;
      console.warn(
        "[ooxml] text selection is unavailable in mode: 'worker'; the overlay will be empty. Use mode: 'main' for selectable text.",
      );
    }
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
  private async _renderSlotBitmap(
    i: number,
    slot: PageSlot,
    widthPx: number,
    dpr: number,
    scale: number,
  ): Promise<void> {
    // Worker-mode + enableTextSelection: the overlay can't be populated (onTextRun
    // doesn't cross the worker boundary), so warn once (parity with DocxViewer)
    // and leave the overlay empty. Fires before the coalescing guards so it is
    // reported even when this particular dispatch is coalesced/dropped.
    this._maybeWarnNoTextSelection();
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
      // This bitmap now defines the scale the on-screen canvas lives at, so a
      // later zoom preview stretches from HERE (design §7 renderedScale).
      slot.renderedScale = scale;
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
        void this._renderSlotBitmap(i, live, this._pageWidthPx(i), this._dpr(), this._scale);
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
   * FLICKER-FREE (design §7): this does NOT re-render the visible pages inline.
   * It shows an immediate CSS preview (stretch the existing bitmaps, scale the
   * overlays) and DEBOUNCES a full-resolution settle re-render for ZOOM_SETTLE_MS,
   * so a wheel/pinch burst never blanks a page and coalesces into one crisp render.
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
    const r1 = computeVisibleRange(
      this._heights,
      this._gap(),
      0,
      this._scrollHost.clientHeight,
      this._overscan(),
      this._pad(),
    );
    this._spacer.style.height = `${r1.totalHeight}px`;
    // The page px width changed with the scale, so the horizontal extent moves too.
    this._syncSpacerWidth();

    // Pin the same fractional position of the same page under the viewport top.
    const maxTop = Math.max(0, r1.totalHeight - this._scrollHost.clientHeight);
    const wantTop = (r1.offsets[top] ?? 0) + intraFrac * (this._heights[top] || 0);
    this._scrollHost.scrollTop = Math.min(maxTop, Math.max(0, wantTop));

    // FLICKER-FREE ZOOM (design §7). Do NOT recycle + re-render in-window slots
    // (that blanks each visible page to white every tick). Instead:
    //  1. CSS-PREVIEW the currently-mounted slots at the new geometry — reposition
    //     the wrapper, stretch the existing canvas bitmap via style.width/height
    //     (soft but never blank), and scale the text overlay by the ratio between
    //     the new scale and the scale the overlay was built at.
    //  2. DEBOUNCE a full-resolution settle re-render: schedule it ZOOM_SETTLE_MS
    //     after the LAST setScale so a wheel/pinch burst coalesces into one render.
    this._previewVisible();
    this._scheduleSettle();
  }

  /**
   * CSS preview of the visible window at the current `_scale` (design §7
   * mechanism 1), WITHOUT re-rendering. Slots leaving the window recycle normally;
   * slots ENTERING the window mount fresh (rendered at the current scale directly,
   * so they never need a preview); slots that STAY are repositioned and their
   * canvas + text overlay are CSS-transformed to the new size (the device buffer
   * is untouched — that is the whole point: no synchronous clear, no blank frame).
   */
  private _previewVisible(): void {
    if (!this._doc || this._doc.pageCount === 0) return;
    const r = this._range();
    this._lastRange = r;

    // Recycle slots that left [start, end].
    for (const [idx, slot] of [...this._slots]) {
      if (idx < r.start || idx > r.end) this._recycleSlot(idx, slot);
    }
    // For every index in the window: mount fresh if missing (renders at the current
    // scale), or CSS-preview if already mounted (no re-render, no device resize).
    for (let i = r.start; i <= r.end; i++) {
      const existing = this._slots.get(i);
      if (!existing) {
        const slot = this._acquireSlot();
        this._positionSlot(slot, i, r);
        this._slots.set(i, slot);
        this._renderSlot(i, slot);
      } else {
        this._previewSlot(existing, i, r);
      }
    }
    // Fire onVisiblePageChange only when the top page actually changed.
    if (r.topIndex !== this._lastTopIndex) {
      this._lastTopIndex = r.topIndex;
      this._opts.onVisiblePageChange?.(r.topIndex, this._doc.pageCount);
    }
  }

  /**
   * CSS-preview a single already-mounted slot at the new geometry (design §7): the
   * wrapper is repositioned + sized (via `_positionSlot`), the canvas bitmap is
   * STRETCHED to the new CSS size (no `canvas.width` — the device buffer, and thus
   * the drawn pixels, are left intact, just scaled by the browser), and the text
   * overlay is scaled by `newScale / renderedScale` so it tracks the stretched
   * page. `renderedScale <= 0` means the slot's first render hasn't resolved yet
   * (nothing to stretch); the pending render captured the current scale, so it
   * lands correct and no preview is needed.
   */
  private _previewSlot(slot: PageSlot, i: number, r: VisibleRange): void {
    this._positionSlot(slot, i, r);
    // Stretch the existing bitmap to the new CSS box (device buffer untouched).
    slot.canvas.style.width = `${this._pageWidthPx(i)}px`;
    slot.canvas.style.height = `${this._pageHeightPx(i)}px`;
    if (slot.textLayer && slot.renderedScale > 0) {
      const ratio = this._scale / slot.renderedScale;
      slot.textLayer.style.transformOrigin = '0 0';
      slot.textLayer.style.transform = `scale(${ratio})`;
    }
  }

  /** (Re)schedule the debounced settle re-render (design §7 mechanism 2). Resets
   *  the timer on every call so a burst of `setScale` dispatches ONE settle
   *  ZOOM_SETTLE_MS after the LAST call. Cleared in `destroy()`. */
  private _scheduleSettle(): void {
    if (this._settleTimer !== null) clearTimeout(this._settleTimer);
    this._settleTimer = setTimeout(() => {
      this._settleTimer = null;
      this._settleRender();
    }, ZOOM_SETTLE_MS);
  }

  /** Full-resolution settle re-render of the visible window (design §7 mechanisms
   *  2+3). Re-renders each mounted slot at the current scale via the double-buffer
   *  swap (main) / same-canvas transfer (worker). Main mode also rebuilds the text
   *  overlay and clears its preview transform; in worker mode the overlay is
   *  permanently empty (text selection is main-mode-only), so the transform is
   *  inert there and is reset on recycle. Dispatched at the CURRENT epoch; the
   *  existing epoch gate discards it if a later `setScale` supersedes it
   *  mid-render. */
  private _settleRender(): void {
    if (this._destroyed || !this._doc || this._doc.pageCount === 0) return;
    for (const [i, slot] of [...this._slots]) {
      // Skip slots already at the current scale (a slot that entered the window
      // during the burst mounted fresh at the current scale — nothing to settle).
      if (slot.renderedScale === this._scale) continue;
      this._settleSlot(i, slot);
    }
  }

  /**
   * Settle-render one slot at the current scale (design §7 mechanism 3).
   *
   * WORKER: re-dispatch the bitmap render into the SAME canvas. The worker path
   * sizes the device buffer and `transferFromImageBitmap`s it in ONE synchronous
   * step (no await between `canvas.width = …` and the transfer), so the browser
   * never composites an intermediate blank frame — no spare canvas is needed. The
   * `renderedScale === _scale` gate in `_settleRender` plus the epoch gate inside
   * `_renderSlotBitmap` keep this correct and idempotent.
   *
   * MAIN: `renderPage` (via renderDocumentToCanvas) synchronously sets
   * `canvas.width = …` (which CLEARS the backing store to blank) BEFORE its first
   * await and paints AFTER — so rendering into the on-screen canvas would flash it
   * white. Render into a SPARE off-DOM canvas instead; only once it resolves at the
   * current epoch do we swap it into the wrapper (replacing the old canvas, which is
   * DISCARDED — the pooled unit is the slot, not the canvas). The old canvas keeps
   * showing the stretched preview until the instant of the swap — blank-free.
   */
  private _settleSlot(i: number, slot: PageSlot): void {
    if (!this._doc) return;
    const dpr = this._dpr();
    const widthPx = this._pageWidthPx(i);
    const scale = this._scale;
    const epoch = this._renderEpoch;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr, scale);
      return;
    }

    // Main mode: double-buffer. Render into a spare canvas kept off-DOM. The
    // spare REPLACES the on-screen canvas on swap, so it must carry the page
    // shadow too — otherwise a settle would silently drop it.
    const spare = document.createElement('canvas');
    spare.style.cssText = 'display:block;background:#fff;';
    this._applyPageShadow(spare);
    const runs: DocxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const onTextRun = wantOverlay ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
    this._doc
      .renderPage(spare, i, {
        width: widthPx,
        dpr,
        defaultTextColor: this._opts.defaultTextColor,
        showTrackChanges: this._opts.showTrackChanges,
        onTextRun,
      })
      .then(() => {
        // Discard if superseded: a later setScale bumped the epoch (this spare is
        // at a stale scale), or the slot recycled / moved to another page. Drop
        // the spare (it is off-DOM, so GC reclaims it) and do NOT swap.
        if (epoch !== this._renderEpoch || this._slots.get(i) !== slot || slot.renderedPage !== i) return;
        // Swap the freshly-painted spare in for the old (stretched-preview) canvas.
        // The old canvas was the only child that showed content; replacing it in
        // one DOM op means the screen goes from preview → crisp with no blank tick.
        const old = slot.canvas;
        slot.wrapper.insertBefore(spare, old);
        old.remove();
        slot.canvas = spare;
        // The retired canvas held a 2d context; keep the pool clean by dropping any
        // bitmaprenderer handle association (main-mode canvases never had one).
        slot.bitmapCtx = null;
        slot.renderedScale = scale;
        // Rebuild the overlay at the full resolution and CLEAR the preview
        // transform (the crisp render no longer needs the scale()).
        if (slot.textLayer) {
          slot.textLayer.style.transform = '';
          slot.textLayer.style.transformOrigin = '';
          if (wantOverlay) {
            buildDocxTextLayer(
              slot.textLayer,
              runs,
              spare.style.width || `${spare.width}px`,
              spare.style.height || `${spare.height}px`,
            );
          }
        }
      })
      .catch((err: unknown) => {
        this._reportRenderError(err);
      });
  }

  /**
   * Scroll so page `index`'s top edge sits at the viewport top. Clamps `index` to
   * `[0, pageCount-1]` (the pager convention) and the resulting scrollTop to
   * `[0, totalHeight − viewportHeight]` so the last pages don't scroll past the
   * end. A no-op when nothing is loaded or the document is empty.
   *
   * `opts.behavior` ('auto' | 'smooth', default 'auto') is honoured via
   * `scrollHost.scrollTo({ top, behavior })` when the host supports it (a real
   * browser); the stub-DOM has no `scrollTo`, so the fallback sets `scrollTop`
   * directly (which is what the tests assert). We then call `_mountVisible` once.
   *
   * MOUNTING CAVEAT: synchronous mounting of the target page is guaranteed only on
   * the DEFAULT/'auto' path — there `scrollTop` has already jumped to `top`, so the
   * `_mountVisible` call reads the final scroll position and the target page's slots
   * exist immediately. With `behavior: 'smooth'` the scroll animates ASYNCHRONOUSLY:
   * `scrollTop` is still near the old position when `_mountVisible` runs, so the
   * target page mounts lazily via the animation's subsequent `scroll` events, not
   * from this call.
   */
  scrollToPage(index: number, opts?: { behavior?: 'auto' | 'smooth' }): void {
    if (!this._doc || this._doc.pageCount === 0 || !this._scaleEstablished) return;
    const clamped = Math.max(0, Math.min(index, this._doc.pageCount - 1));
    // Recompute offsets from the current heights (independent of scrollTop).
    const r = computeVisibleRange(
      this._heights,
      this._gap(),
      0,
      this._scrollHost.clientHeight,
      this._overscan(),
      this._pad(),
    );
    const target = r.offsets[clamped] ?? 0;
    const maxTop = Math.max(0, r.totalHeight - this._scrollHost.clientHeight);
    const top = Math.min(maxTop, Math.max(0, target));
    const host = this._scrollHost as HTMLDivElement & {
      scrollTo?: (opts: { top: number; behavior?: 'auto' | 'smooth' }) => void;
    };
    if (typeof host.scrollTo === 'function') {
      host.scrollTo({ top, behavior: opts?.behavior ?? 'auto' });
    } else {
      this._scrollHost.scrollTop = top;
    }
    this._mountVisible();
  }

  /**
   * Re-fit the base scale on a container resize while PRESERVING the current zoom
   * multiplier (design §11), then re-anchor + re-render. A `ResizeObserver` fires
   * on any box change, but only a WIDTH change alters the fit-to-width base scale;
   * a height-only change skips the re-fit yet STILL re-mounts the visible window
   * (via `_mountVisible`), because a taller viewport reveals rows that were below
   * the fold and would otherwise stay blank until the next scroll. Empty/unloaded
   * ⇒ no-op; a still-zero width ⇒ defer.
   *
   * Zero-width recovery: a container that was 0-wide at construction never
   * established a scale (`_scaleEstablished` is false), so the first non-zero
   * resize establishes it here via `relayout()` — completing the T2 deferral.
   *
   * Re-fit math (zoom multiplier preserved):
   *   mult      = _scale / _prevBase            (the user's zoom over the old base)
   *   newScale  = newBase × mult
   * Routing through `setScale(newScale)` bumps `_renderEpoch` (resize IS an epoch
   * event — T4 banner) and re-anchors + CSS-previews + debounces a settle re-render
   * of every slot at the new geometry, exactly like a zoom (design §7 flicker-free
   * path — a rapid ResizeObserver burst therefore also coalesces into one settle).
   * `setScale`'s clamp/no-op guards apply: an unchanged newScale (identical width)
   * is a no-op there — so we short-circuit BEFORE it when the fit-width is
   * unchanged (mounting the revealed window without a needless re-render), and
   * after it we call `_mountVisible` again to cover the case where the clamp made
   * `setScale` no-op yet the viewport still grew.
   */
  private _onResize(): void {
    if (!this._doc || this._doc.pageCount === 0) return;
    // Zero-width recovery: first non-zero layout establishes the base scale.
    if (!this._scaleEstablished) {
      this.relayout();
      return;
    }
    const newBase = this._baseScale();
    if (newBase <= 0) return; // still unlaid-out — wait for the next resize
    const newFitWidth = this._fitWidthPx();
    if (newFitWidth === this._lastFitWidth) {
      // Height-only change (or any resize that leaves the fit-width identical):
      // the base scale is unchanged, so there is no re-fit to do — but a taller
      // viewport now exposes rows that were below the fold. `_mountVisible`
      // recomputes the visible range from the CURRENT clientHeight and mounts the
      // newly-revealed pages; without it those rows stay blank until the user
      // scrolls (which recomputes the range). No epoch bump — the geometry
      // (and every mounted slot's px size) is unchanged, so cached canvases are
      // still valid; we only add the missing slots.
      this._mountVisible();
      return;
    }
    this._lastFitWidth = newFitWidth;
    // Preserve the zoom multiplier across the re-fit: newScale = newBase × mult.
    const mult = this._prevBase > 0 ? this._scale / this._prevBase : 1;
    this._prevBase = newBase;
    // Route through setScale so the epoch bumps and the re-anchor/force-re-render
    // path runs identically to a zoom.
    //
    // zoomMin RATCHET (design §8.1 caveat, see setScale JSDoc): `zoomMin`/`zoomMax`
    // are ABSOLUTE px-per-pt bounds, but the re-fit base (`newBase × mult`) is
    // computed UNCLAMPED. A resize that transits the scale below `zoomMin × pageWidth`
    // (a wide page in a container that briefly narrows) is clamped UP by `setScale`,
    // which permanently inflates the implied multiplier even with zero user zoom —
    // the next re-fit reads back the clamped `_scale` as `mult`. This is bounded and
    // converges (the clamp floor is fixed), but it means the preserved multiplier can
    // drift above 1 purely from resize transits below the floor. Accepted consequence
    // of using absolute bounds (§8.1) with an unclamped relayout base.
    this.setScale(newBase * mult);
    // `setScale` no-ops when the clamped scale is unchanged (e.g. already pinned at
    // a clamp boundary), which would skip its preview + settle. A width+height
    // growth that ends up clamped to the same scale must still reveal the taller
    // viewport's rows, so mount here too. Idempotent when `setScale` ran: the
    // window is already mounted and every present slot is a re-position no-op.
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

  /** @internal test hook: the current render epoch (bumped on setScale + resize). */
  renderEpochForTest(): number {
    return this._renderEpoch;
  }

  /** @internal test hook: fire the observed resize path (a real host drives this
   *  via the constructor's ResizeObserver). */
  resizeForTest(): void {
    this._onResize();
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
    this._resizeObserver?.disconnect();
    this._resizeObserver = null;
    // Cancel a pending settle so no re-render is dispatched after teardown
    // (design §7 mechanism 2). `_destroyed` also guards `_settleRender`, but
    // clearing the timer avoids the wasted wake-up and keeps fake-timer tests
    // deterministic.
    if (this._settleTimer !== null) {
      clearTimeout(this._settleTimer);
      this._settleTimer = null;
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
