import { computeVisibleRange, EMU_PER_PX, zoomStepScale, type VisibleRange } from '@silurus/ooxml-core';
import { PptxPresentation, type LoadOptions, type RenderSlideOptions } from './presentation';
import type { PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';

/**
 * Options for {@link PptxScrollViewer}. Extends `RenderSlideOptions` (per-slide
 * render knobs, minus `onTextRun`) and `LoadOptions` (parse/worker knobs). See
 * design §8.2.
 *
 * `onTextRun` is omitted deliberately: the viewer drives it internally per
 * mounted slot to build the optional per-slide selection overlay (gated by
 * `enableTextSelection`), so exposing it here would let a caller's callback be
 * silently overridden.
 *
 * NOTE: `RenderSlideOptions` also carries `dim` and `skipMediaControls`. The v1
 * scroll viewer never sets `dim` or `skipMediaControls` (hidden-slide dimming is
 * a PAGER policy, not a scroll-viewer feature — design §8.2 / Delta 6). These
 * inherited fields are accepted for type-compatibility but are not part of the
 * scroll-viewer's supported API.
 */
export interface PptxScrollViewerOptions extends Omit<RenderSlideOptions, 'onTextRun'>, LoadOptions {
  /** Base fit width in CSS px → base zoom scale. Default: the container's width
   *  at first non-zero layout (design §7/§11 zero-width deferral). */
  width?: number;
  /** Vertical gap (px) between consecutive slides. Default 16. */
  gap?: number;
  /** Desk padding (px) ABOVE the FIRST slide — the margin a presentation viewer
   *  leaves between the top of the scroll surface and the first slide. Default:
   *  `gap` (uniform desk rhythm — the first slide sits the same distance from the
   *  top as slides sit from each other). Pass `0` for a flush-top layout. */
  paddingTop?: number;
  /** Desk padding (px) BELOW the LAST slide — the margin below the final slide.
   *  Default: `gap`. Pass `0` for a flush-bottom layout. */
  paddingBottom?: number;
  /** Slides kept mounted beyond the viewport on each side. Default 1. */
  overscan?: number;
  /** Per-slide transparent text-selection overlay. MAIN render mode only:
   *  in worker mode `onTextRun` cannot cross the worker boundary, so the overlay
   *  stays empty and the viewer logs one warning (design §11). */
  enableTextSelection?: boolean;
  /** Minimum zoom scale — a DIMENSIONLESS multiplier over the 96-dpi natural
   *  slide size (10% = 0.1), matching `DocxScrollViewer`. Default 0.1. */
  zoomMin?: number;
  /** Maximum zoom scale (dimensionless multiplier, 400% = 4). Default 4. */
  zoomMax?: number;
  /** Enable `Ctrl`/`Cmd`+wheel zoom. Default true. */
  enableZoom?: boolean;
  /**
   * CSS `background` shorthand for the scroll surface (the "desk") visible
   * behind and between slides — the gray a presentation viewer paints around the
   * slide. Applied to the viewer-owned scroll host. The slides themselves are
   * always drawn on their own white canvas and are unaffected. Default
   * `undefined`: the scroll surface stays transparent so the host container's
   * background shows through (non-breaking).
   */
  background?: string;
  /**
   * Inject an already-loaded engine to share one parse across panes (design §14).
   * When set: `load()` is unsupported (throws), the engine's own `mode` wins (an
   * explicitly conflicting `opts.mode` throws at construction, design §11), and
   * `destroy()` does NOT destroy this engine (the caller owns its lifecycle).
   */
  presentation?: PptxPresentation;
  /** Fires when the top-most visible slide changes. `topIndex` from
   *  `computeVisibleRange` (the first slide intersecting the viewport top,
   *  EXCLUDING overscan). */
  onVisibleSlideChange?: (topIndex: number, total: number) => void;
  /** Error callback. When set, `load()` invokes it and resolves (otherwise the
   *  error is rethrown — shared viewer error contract). It ALSO fires for async
   *  per-slot render failures (both main `renderSlide` and worker
   *  `renderSlideToBitmap` rejections); a failed slide is left blank rather than
   *  crashing the loop. Without an `onError`, render failures are logged via
   *  `console.error` so they are never fully silent. */
  onError?: (err: Error) => void;
}

/** One mounted slide. `canvas` is the drawn slide; `textLayer` the optional
 *  per-slide selection overlay (main mode only). `renderedSlide` guards against
 *  re-rendering a recycled slot for a slide whose render is still in flight. */
interface SlideSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  /** slide index this slot is currently rendering / has rendered, or -1 when free. */
  renderedSlide: number;
  /** worker-mode: a transient hold on a just-received ImageBitmap, set only
   *  between receipt from the worker and its `transferFromImageBitmap` (which
   *  consumes it, after which we null the field). Its purpose is the throw path:
   *  if the transfer throws, `destroy()`/`_recycleSlot` can still find and
   *  `.close()` the bitmap. Normally null once transfer completes. */
  bitmap: ImageBitmap | null;
  /** bitmaprenderer ctx (worker mode), grabbed once per canvas. */
  bitmapCtx: ImageBitmapRenderingContext | null;
}

export class PptxScrollViewer {
  private _pres: PptxPresentation | null = null;
  private readonly _injected: boolean;
  private readonly _opts: PptxScrollViewerOptions;
  private readonly _container: HTMLElement;
  private readonly _wrapper: HTMLDivElement;
  private readonly _scrollHost: HTMLDivElement;
  private readonly _spacer: HTMLDivElement;
  /** Resolved render mode. When an engine is injected the engine's own `mode`
   *  is authoritative (design §11 — no silent mis-pathing / no probing); an
   *  explicitly conflicting `opts.mode` is rejected at construction. When self-
   *  loading, `opts.mode` decides and `load()` passes it to `PptxPresentation.load`. */
  private _mode: 'main' | 'worker';

  /** Dimensionless zoom multiplier over the 96-dpi natural slide size (mirrors
   *  `DocxScrollViewer`, whose `_scale` multiplies `widthPt × PT_TO_PX`). The
   *  natural (1×) slide width in CSS px is `slideEmu / EMU_PER_PX`; the base fit
   *  sets `_scale` so that natural width maps to the container width, and zoom
   *  multiplies it further (design §7). */
  private _scale = 1;
  /** Whether the base fit scale has been established. Set true the first time
   *  `relayout()` resolves a positive base scale. We use an explicit flag rather
   *  than a `_scale === 1` sentinel because a fit scale of exactly 1 is a valid
   *  established state (a 1× fit would otherwise be re-fit forever). */
  private _scaleEstablished = false;
  /** Live slots keyed by slide index. */
  private readonly _slots = new Map<number, SlideSlot>();
  /** Recyclable detached slots (canvas + textLayer reused across slides). */
  private readonly _free: SlideSlot[] = [];
  /** Cached per-slide heights in px at the current scale (index-aligned). All
   *  slides are the same size, so every entry equals the uniform slide height. */
  private _heights: number[] = [];
  private _lastRange: VisibleRange | null = null;
  private _lastTopIndex = -1;
  private _scrollListener: (() => void) | null = null;
  /** Set by `destroy()`. Async render callbacks (main + worker) check it before
   *  reporting an error so a rejection that lands after teardown is swallowed
   *  rather than surfaced to a `onError` on a dead viewer. */
  private _destroyed = false;
  /** Worker mode: slide indices whose bitmap render is currently dispatched to the
   *  engine. Coalesces a scroll storm — we never dispatch a second render for a
   *  slide whose first is still in flight — and lets us drop slides that scrolled
   *  out of the window before dispatch (design §11 worker coalescing).
   *
   *  T4 ZOOM HAZARD (RESOLVED by the render epoch below): coalescing keys on slide
   *  INDEX only, with no notion of the scale a dispatch was made at. Once
   *  `setScale` can change the zoom mid-flight, an in-flight bitmap dispatched at
   *  the OLD scale can still pass the on-resolution identity check if the SAME
   *  slot object is re-mounted for slide `i` (the pool reuses slot objects, so
   *  `_slots.get(i) === slot && slot.renderedSlide === i` can hold for an old
   *  dispatch), and get painted at the WRONG resolution. We fix this with a render
   *  epoch (`_renderEpoch`): each dispatch captures the epoch, and on resolution a
   *  moved epoch ⇒ STALE (close + re-dispatch the live slot). See
   *  `_renderSlotBitmap`. */
  private readonly _slideInFlight = new Set<number>();
  /** Render generation, bumped on every effective `setScale` (and the resize
   *  re-fit in `_onResize`, which routes through `setScale`). Stamped into each async render
   *  dispatch; a resolution whose captured epoch ≠ this value is STALE — its
   *  pixels/geometry are at a superseded scale. Worker path: close the orphan
   *  bitmap + re-dispatch the live slot. Main path: skip the (stale) text-layer
   *  build; the engine's per-canvas token already discards the stale pixels. */
  private _renderEpoch = 0;
  private _wheelListener: ((e: WheelEvent) => void) | null = null;
  /** One-shot latch for the worker-mode text-selection warning. The overlay is a
   *  main-mode-only feature: in worker mode the per-run `onTextRun` geometry
   *  cannot cross the worker boundary, so an `enableTextSelection` overlay stays
   *  empty. We warn once (parity with `PptxViewer`) rather than per slot. */
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

  constructor(container: HTMLElement, opts: PptxScrollViewerOptions = {}) {
    this._container = container;
    this._opts = opts;
    this._injected = !!opts.presentation;
    if (this._injected) {
      const engine = opts.presentation as PptxPresentation;
      // Injected engine ⇒ its own mode is the fact (design §11). An EXPLICITLY
      // conflicting opts.mode is a mis-configuration and is rejected here; an
      // absent opts.mode is fine.
      if (opts.mode !== undefined && opts.mode !== engine.mode) {
        throw new Error(
          `PptxScrollViewer: opts.mode='${opts.mode}' conflicts with the injected engine's mode='${engine.mode}'. ` +
            'Omit opts.mode when injecting an engine — the engine owns its render mode.',
        );
      }
      this._pres = engine;
      this._mode = engine.mode;
    } else {
      this._mode = opts.mode ?? 'main';
    }

    // container → wrapper → scrollHost → spacer  (design §6)
    this._wrapper = document.createElement('div');
    this._wrapper.style.cssText = 'position:relative;width:100%;height:100%;overflow:hidden;';
    this._scrollHost = document.createElement('div');
    this._scrollHost.style.cssText = 'position:absolute;inset:0;overflow:auto;';
    // The "desk" behind/between slides. Undefined ⇒ transparent (container shows
    // through); slides keep their own white canvas regardless.
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
   * Load a PPTX from URL or ArrayBuffer and render the first window.
   * UNSUPPORTED when an engine was injected via `opts.presentation` (throws) — the
   * caller already owns the parsed engine.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    if (this._injected) {
      throw new Error(
        'PptxScrollViewer.load() is unsupported when an engine is injected via opts.presentation; the injected engine is already loaded.',
      );
    }
    try {
      this._pres = await PptxPresentation.load(source, {
        useGoogleFonts: this._opts.useGoogleFonts,
        maxZipEntryBytes: this._opts.maxZipEntryBytes,
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

  get slideCount(): number {
    return this._pres?.slideCount ?? 0;
  }

  /** Uniform slide width in CSS px at the current scale. `_scale` is a
   *  dimensionless multiplier over the natural 96-dpi width (`slideEmu /
   *  EMU_PER_PX`), mirroring docx's `widthPt × PT_TO_PX × _scale`. */
  private _slideWidthPx(): number {
    return (this._pres!.slideWidth / EMU_PER_PX) * this._scale;
  }

  /** Uniform slide height in CSS px at the current scale. */
  private _slideHeightPx(): number {
    return (this._pres!.slideHeight / EMU_PER_PX) * this._scale;
  }

  /** The container (fit) width, deferring when the container is unlaid-out. */
  private _fitWidthPx(): number {
    if (this._opts.width && this._opts.width > 0) return this._opts.width;
    const cw = this._container.clientWidth || this._scrollHost.clientWidth;
    return cw > 0 ? cw : 0; // 0 ⇒ defer (design §11 zero-width deferral)
  }

  /** Base scale: the DIMENSIONLESS multiplier that fits the (uniform) slide
   *  width to the fit-width. `natural = slideWidthEmu / EMU_PER_PX` is the 96-dpi
   *  CSS-px width; `base = fitWidth / natural` (mirrors docx's `w / (widthPt ×
   *  PT_TO_PX)`). Returns 0 when the container has no width yet (deferral). */
  private _baseScale(): number {
    if (!this._pres || this._pres.slideCount === 0) return 0;
    const w = this._fitWidthPx();
    const naturalW = this._pres.slideWidth / EMU_PER_PX;
    if (w <= 0 || naturalW <= 0) return 0;
    return w / naturalW; // dimensionless multiplier over the natural width
  }

  /**
   * Recompute per-slide heights + the spacer and re-mount the visible window.
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
    if (!this._pres) return;
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

  /** All slides are the same size, so heights = n × uniform. We still feed this
   *  full array to computeVisibleRange (never special-case uniform) so offsets /
   *  topIndex live in one tested place (design §5.1). */
  private _recomputeHeights(): void {
    const n = this._pres!.slideCount;
    const h = this._slideHeightPx();
    this._heights = new Array<number>(n).fill(h);
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
  }

  private _onScroll(): void {
    if (!this._pres || !this._scaleEstablished) return;
    this._mountVisible();
  }

  /** Mount/recycle slots for the current visible window. */
  private _mountVisible(): void {
    if (!this._pres || this._pres.slideCount === 0) return;
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
    // onVisibleSlideChange fires ONLY when the top visible slide actually changes
    // (change-only latch; `_lastTopIndex` starts at -1 so the first layout fires
    // once for slide 0). Every mount path — scroll, zoom, resize re-fit, and
    // scrollToSlide — funnels through here, so navigation never double-fires.
    if (r.topIndex !== this._lastTopIndex) {
      this._lastTopIndex = r.topIndex;
      this._opts.onVisibleSlideChange?.(r.topIndex, this._pres.slideCount);
    }
  }

  private _acquireSlot(): SlideSlot {
    const reused = this._free.pop();
    if (reused) {
      // _recycleSlot already reset renderedSlide to -1 before pooling this slot.
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
    const slot: SlideSlot = { wrapper, canvas, textLayer, renderedSlide: -1, bitmap: null, bitmapCtx: null };
    return slot;
  }

  private _recycleSlot(idx: number, slot: SlideSlot): void {
    this._slots.delete(idx);
    // Close any worker bitmap held by this slot (T3 sets slot.bitmap).
    if (slot.bitmap) {
      slot.bitmap.close();
      slot.bitmap = null;
    }
    // Clear the per-slot text overlay so a slot sitting in the free pool holds no
    // stale spans. buildPptxTextLayer also clears on its next build, but an
    // unrendered pooled slot never gets that build, and the detached spans would
    // otherwise linger; drop them here.
    if (slot.textLayer) slot.textLayer.innerHTML = '';
    slot.renderedSlide = -1;
    slot.wrapper.remove();
    this._free.push(slot);
  }

  private _positionSlot(slot: SlideSlot, i: number, r: VisibleRange): void {
    slot.wrapper.style.top = `${r.offsets[i]}px`;
    slot.wrapper.style.width = `${this._slideWidthPx()}px`;
    slot.wrapper.style.height = `${this._slideHeightPx()}px`;
  }

  /** Device-pixel ratio for a render (opts override → window → 1). */
  private _dpr(): number {
    return this._opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
  }

  /**
   * Render slide `i` into `slot`. Routes strictly on the constructor-resolved
   * `_mode` (design §11 — no probing, no silent mis-pathing): `main` ⇒ paint the
   * slot's canvas directly via `renderSlide`; `worker` ⇒ transfer an ImageBitmap
   * from `renderSlideToBitmap`.
   *
   * Slot-identity guard: a slot recycled to a DIFFERENT slide while a previous
   * render is in flight must not repaint the stale slide. `slot.renderedSlide`
   * tracks the slide this slot is committed to; we stamp it up-front and bail on
   * resolution if it changed (the engine's own token guard is per-canvas; this is
   * the viewer's per-slot slide-identity check).
   *
   * Render epoch (main path): pixel staleness after a mid-flight `setScale` is
   * already handled by the engine's per-canvas token (the newer renderSlide on the
   * same canvas wins) — `setScale` recycles + re-mounts, and the re-mount always
   * re-dispatches `renderSlide` (renderedSlide reset to -1), so a fresh render is
   * always issued. But the viewer-side side effects of a STALE resolution — the
   * text-layer build (its run geometry is at the OLD scale) and the renderedSlide
   * bookkeeping — must NOT run, or a superseded render would rebuild the overlay
   * with stale x/y/w/h (the pool reuses slot objects, so the identity check alone
   * can pass for an old-epoch resolution). We gate them on the captured epoch.
   */
  private _renderSlot(i: number, slot: SlideSlot): void {
    if (!this._pres) return;
    // Slot-identity guard: this slot is already rendering / has rendered slide i.
    if (slot.renderedSlide === i) return;
    slot.renderedSlide = i;

    const dpr = this._dpr();
    const widthPx = this._slideWidthPx();
    const epoch = this._renderEpoch;

    if (this._mode === 'worker') {
      void this._renderSlotBitmap(i, slot, widthPx, dpr);
      return;
    }

    // Main mode: render straight onto the slot's canvas.
    const runs: PptxTextRunInfo[] = [];
    const wantOverlay = !!this._opts.enableTextSelection && !!slot.textLayer;
    const onTextRun = wantOverlay ? (r: PptxTextRunInfo) => runs.push(r) : undefined;
    this._pres
      .renderSlide(slot.canvas, i, {
        width: widthPx, // this slide's own px width → uniform px-per-EMU scale (§7)
        dpr,
        onTextRun,
      })
      .then(() => {
        // Stale if the epoch moved (a setScale rescaled mid-flight — the run
        // geometry is at the old scale), or a recycle re-purposed this slot for a
        // different slide / freed it. Either way: skip the (stale) overlay build.
        // The engine's per-canvas token already discards the superseded pixels.
        if (epoch !== this._renderEpoch || this._slots.get(i) !== slot || slot.renderedSlide !== i) return;
        if (wantOverlay && slot.textLayer) {
          // buildPptxTextLayer takes NUMBERS (not strings) for width/height. The
          // overlay must match the slot's CSS box, NOT the canvas backing store:
          // renderSlide sets `canvas.width = cssWidth × dpr`, so on a retina (dpr 2)
          // display the backing store is 2× the CSS box. Passing it would size the
          // overlay 2× too large (overflowing the wrapper + inflating the scroll
          // area). Pass the CSS px directly — the uniform slide width/height at the
          // current scale (rounded).
          buildPptxTextLayer(slot.textLayer, runs, Math.round(widthPx), Math.round(this._slideHeightPx()));
        }
      })
      .catch((err: unknown) => {
        this._reportRenderError(err);
      });
  }

  /** Warn once when an `enableTextSelection` overlay was requested but the render
   *  mode is `worker` (so the overlay stays empty). Same wording as
   *  `PptxViewer` — one warning per viewer, not per slot. */
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
    else console.error('[ooxml] PptxScrollViewer render failed:', e);
  }

  /**
   * Worker-mode slot render: dispatch `renderSlideToBitmap`, transfer the result
   * via a per-slot `bitmaprenderer` context, and manage the ImageBitmap lifecycle.
   *
   * Coalescing / drop-stale (design §11):
   *  - Skip if slide `i` is already in flight (a scroll storm won't double-dispatch).
   *  - Skip if slide `i` already left the mounted window before dispatch.
   *  - On resolution, if `slot` is no longer THIS slide's live slot (it recycled to
   *    another slide, or slide `i` re-mounted onto a DIFFERENT slot while this render
   *    was in flight), close the orphan bitmap and skip the paint. In that
   *    re-mount case a live slot for `i` still awaits a render, so once we clear
   *    the in-flight guard we re-dispatch it — a slide that recycled and re-mounted
   *    mid-flight must never stay blank.
   *  - RENDER EPOCH: the dispatch captures `this._renderEpoch`. `setScale` bumps
   *    the epoch, so a resolution whose captured epoch ≠ the live epoch is STALE
   *    even when the SAME slot object is still mounted for slide `i` (the pool
   *    reuses slot objects, so the identity check alone can't catch a zoom that
   *    happened mid-flight). A moved epoch ⇒ close the orphan + re-dispatch the
   *    live slot at the new scale, never paint the old-scale bitmap.
   *
   * Do NOT pass `dim` or `skipMediaControls` to `renderSlideToBitmap`. The scroll
   * viewer never dims slides (design §8.2 / Delta 6); passing neither means the
   * static play-badge renders on media slides (matching `PptxViewer`'s
   * non-media-playback path) — acceptable for v1.
   */
  private async _renderSlotBitmap(i: number, slot: SlideSlot, widthPx: number, dpr: number): Promise<void> {
    // Worker-mode + enableTextSelection: the overlay can't be populated (onTextRun
    // doesn't cross the worker boundary), so warn once (parity with PptxViewer)
    // and leave the overlay empty. Fires before the coalescing guards so it is
    // reported even when this particular dispatch is coalesced/dropped.
    this._maybeWarnNoTextSelection();
    if (this._slideInFlight.has(i)) return; // coalesce: already dispatched
    // Drop-stale before dispatch: if this slide already scrolled out of the
    // mounted window, don't dispatch at all.
    if (this._slots.get(i) !== slot) return;
    const epoch = this._renderEpoch;
    this._slideInFlight.add(i);
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
      const bmp = await this._pres!.renderSlideToBitmap(i, {
        width: widthPx,
        dpr,
      });
      // Stale if EITHER (a) the epoch moved (a setScale rescaled mid-flight, so
      // this bitmap is at a superseded resolution — this catches the case where
      // the SAME slot object is re-mounted for slide `i`, which the identity check
      // below cannot), or (b) the slot recycled to a different slide / slide `i`
      // re-mounted onto a DIFFERENT slot. Either way: close + skip the paint.
      if (epoch !== this._renderEpoch || this._slots.get(i) !== slot || slot.renderedSlide !== i) {
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
      this._slideInFlight.delete(i);
      // Re-dispatch ONLY when this invocation went stale — a LIVE slot for slide
      // `i` still awaits a correct render and the reason we didn't paint was
      // staleness, not a render failure. The two staleness cases:
      //  - IDENTITY MOVED (`live !== slot`): slide `i` re-mounted onto a DIFFERENT
      //    slot while we ran (the re-mount's own dispatch was coalesced away by
      //    the in-flight guard), so the live slot has no render in flight.
      //  - EPOCH MOVED (`epoch !== this._renderEpoch`): a `setScale` bumped the
      //    epoch mid-flight, so this bitmap was at a superseded scale. The live
      //    slot may be the SAME object reused from the pool, which the identity
      //    test alone would miss — the epoch test catches the same-slot case.
      // NO RETRY ON PLAIN REJECTION: when the slot is still live at the same epoch
      // and we simply failed (`renderSlideToBitmap` rejected or the transfer threw),
      // `!painted` holds but BOTH staleness tests are false, so we do NOT
      // re-dispatch. Retrying a plain failure would loop unbounded (reject →
      // re-dispatch → reject → …); the onError contract is that "a failed slide is
      // left blank" (see PptxScrollViewerOptions.onError), so we leave it blank.
      // Bounded epoch-then-reject: an epoch-moved re-dispatch captures the NEW
      // epoch, so if that fresh render then rejects at the still-current epoch,
      // both tests are false and it stops — no unbounded retry.
      const live = this._slots.get(i);
      if (
        !painted &&
        live &&
        (live !== slot || epoch !== this._renderEpoch) &&
        !this._slideInFlight.has(i) &&
        !this._destroyed
      ) {
        // live.renderedSlide === i already (set by _renderSlot on mount); the fresh
        // dispatch runs at the CURRENT epoch/scale via _slideWidthPx().
        void this._renderSlotBitmap(i, live, this._slideWidthPx(), this._dpr());
      }
    }
  }

  /**
   * Set the absolute (dimensionless) zoom scale — a multiplier over the 96-dpi
   * natural slide size, matching `DocxScrollViewer` — clamped inline to
   * `[zoomMin ?? 0.1, zoomMax ?? 4]` (absolute bounds, XlsxViewer convention — NOT
   * multiples of the base fit; design §3 keeps the clamp in the viewer, not core),
   * then re-anchor VERTICALLY so the slide currently under the viewport top stays
   * fixed. A no-op when nothing is loaded or when the clamped scale is unchanged.
   *
   * Re-anchor (written from scratch — XlsxViewer only re-anchors horizontally):
   * capture `top = topIndex` and the intra-slide fraction `intraFrac` from the
   * CURRENT range BEFORE rescale; after recomputing heights at the new scale,
   * `newScrollTop = offsets'[top] + intraFrac × heights'[top]`, clamped to
   * `[0, totalHeight' − viewportHeight]`. Because a slide's height scales linearly
   * with `_scale`, the same fractional position maps exactly to the new geometry.
   *
   * CAVEAT — base fit below the floor: `relayout()` sets `_scale = base` WITHOUT
   * clamping to `[zoomMin, zoomMax]`. If the base fit is below `zoomMin` (a wide
   * slide in a narrow container), the initial scale sits under the floor, but once
   * the user zooms via `setScale` the clamp pins the minimum to `zoomMin`, so they
   * can no longer return below the floor to the original base fit through this API.
   */
  setScale(scale: number): void {
    if (!this._pres || this._pres.slideCount === 0 || !this._scaleEstablished) return;
    const zoomMin = this._opts.zoomMin ?? 0.1;
    const zoomMax = this._opts.zoomMax ?? 4;
    const next = Math.min(zoomMax, Math.max(zoomMin, scale));
    if (next === this._scale) return;

    // Capture the anchor from the CURRENT layout, before rescale.
    const r0 = this._range();
    const top = r0.topIndex;
    const h0 = this._heights[top] || 0;
    // intraFrac: the fraction of slide `top` that has scrolled above the viewport
    // top. Clamp to [0,1] — a scrollTop inside the trailing gap after slide `top`
    // is attributed to `top` by computeVisibleRange and would push intraFrac past
    // 1, which would drift the re-anchor into the gap; pin the slide instead.
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

    // Pin the same fractional position of the same slide under the viewport top.
    const maxTop = Math.max(0, r1.totalHeight - this._scrollHost.clientHeight);
    const wantTop = (r1.offsets[top] ?? 0) + intraFrac * (this._heights[top] || 0);
    this._scrollHost.scrollTop = Math.min(maxTop, Math.max(0, wantTop));

    // Re-mount at the new geometry, forcing a re-render of every slot (its px
    // size changed under the new scale, so the cached canvas is stale).
    this._mountVisibleForceRerender();
  }

  /** Like `_mountVisible` but recycles every live slot first so each re-mounts and
   *  re-renders at the current scale. Used by `setScale` (and the resize re-fit
   *  in `_onResize`, which routes through `setScale`): a slot's canvas px size
   *  changes with the scale, so the previously
   *  drawn pixels are stale and every visible slide must be redrawn. */
  private _mountVisibleForceRerender(): void {
    for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
    this._mountVisible();
  }

  /**
   * Scroll so slide `index`'s top edge sits at the viewport top. Clamps `index` to
   * `[0, slideCount-1]` (the pager convention) and the resulting scrollTop to
   * `[0, totalHeight − viewportHeight]` so the last slides don't scroll past the
   * end. A no-op when nothing is loaded or the deck is empty.
   *
   * `opts.behavior` ('auto' | 'smooth', default 'auto') is honoured via
   * `scrollHost.scrollTo({ top, behavior })` when the host supports it (a real
   * browser); the stub-DOM has no `scrollTo`, so the fallback sets `scrollTop`
   * directly (which is what the tests assert). We then call `_mountVisible` once.
   *
   * MOUNTING CAVEAT: synchronous mounting of the target slide is guaranteed only on
   * the DEFAULT/'auto' path — there `scrollTop` has already jumped to `top`, so the
   * `_mountVisible` call reads the final scroll position and the target slide's slots
   * exist immediately. With `behavior: 'smooth'` the scroll animates ASYNCHRONOUSLY:
   * `scrollTop` is still near the old position when `_mountVisible` runs, so the
   * target slide mounts lazily via the animation's subsequent `scroll` events, not
   * from this call.
   */
  scrollToSlide(index: number, opts?: { behavior?: 'auto' | 'smooth' }): void {
    if (!this._pres || this._pres.slideCount === 0 || !this._scaleEstablished) return;
    const clamped = Math.max(0, Math.min(index, this._pres.slideCount - 1));
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
   * event — T4 banner) and re-anchors + force-re-renders every slot at the new
   * geometry, exactly like a zoom. `setScale`'s clamp/no-op guards apply: an
   * unchanged newScale (identical width) is a no-op there — so we short-circuit
   * BEFORE it when the fit-width is unchanged (mounting the revealed window without
   * a needless force-re-render), and after it we call `_mountVisible` again to cover
   * the case where the clamp made `setScale` no-op yet the viewport still grew.
   */
  private _onResize(): void {
    if (!this._pres || this._pres.slideCount === 0) return;
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
      // newly-revealed slides; without it those rows stay blank until the user
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
    // zoomMin RATCHET (design §8.2 caveat, see setScale JSDoc): `zoomMin`/`zoomMax`
    // are ABSOLUTE dimensionless bounds, but the re-fit base (`newBase × mult`) is
    // computed UNCLAMPED. A resize that transits the scale below `zoomMin` (a wide
    // slide in a container that briefly narrows) is clamped UP by `setScale`,
    // which permanently inflates the implied multiplier even with zero user zoom —
    // the next re-fit reads back the clamped `_scale` as `mult`. This is bounded and
    // converges (the clamp floor is fixed), but it means the preserved multiplier can
    // drift above 1 purely from resize transits below the floor. Accepted consequence
    // of using absolute bounds (§8.2) with an unclamped relayout base.
    this.setScale(newBase * mult);
    // `setScale` no-ops when the clamped scale is unchanged (e.g. already pinned at
    // a clamp boundary), which would skip its `_mountVisibleForceRerender`. A
    // width+height growth that ends up clamped to the same scale must still reveal
    // the taller viewport's rows, so mount here too. Idempotent when `setScale` ran:
    // the window is already mounted and every present slot is a re-position no-op.
    this._mountVisible();
  }

  get topVisibleSlide(): number {
    return this._lastRange?.topIndex ?? 0;
  }

  /** @internal test hook: slide indices currently mounted. */
  mountedSlideIndicesForTest(): number[] {
    return [...this._slots.keys()];
  }

  /** @internal test hook: the current absolute (dimensionless) zoom scale. */
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
    for (const [idx, slot] of [...this._slots]) this._recycleSlot(idx, slot);
    this._free.length = 0;
    if (!this._injected) {
      this._pres?.destroy();
    }
    this._pres = null;
    this._wrapper.remove();
  }
}
