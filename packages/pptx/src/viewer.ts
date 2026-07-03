import type { RenderOptions, PptxTextRunInfo } from './renderer';
import { buildPptxTextLayer } from './text-layer';
import { PptxPresentation, type LoadOptions } from './presentation';
import type { PresentationHandle } from './presentation-handle';
import { nextVisibleIndex, resolveVisibleIndex, countVisible } from './hidden';
import type { DimOptions } from './types';

/** How {@link PptxViewer} presents hidden slides (`<p:sld show="0">`). */
export type HiddenSlideMode = 'show' | 'skip' | 'dim';

/** Default `'dim'` overlay: 60% white (hidden content shows at 40%). */
const DEFAULT_HIDDEN_DIM: DimOptions = { color: '#ffffff', opacity: 0.6 };

export interface PptxViewerOptions extends RenderOptions, LoadOptions {
  /** Called when a slide finishes rendering */
  onSlideChange?: (index: number, total: number) => void;
  /** Called on parse or render errors */
  onError?: (err: Error) => void;
  /**
   * Enable interactive audio/video playback. When true, slides are rendered
   * via {@link PptxPresentation.presentSlide} so media elements become
   * clickable and the viewer draws its own play/pause chrome. When false
   * (default) the viewer renders a static slide with a non-interactive play
   * badge over media posters.
   */
  enableMediaPlayback?: boolean;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on slide content.
   */
  enableTextSelection?: boolean;
  /**
   * How hidden slides (`<p:sld show="0">`, §19.3.1.38) are presented:
   * - `'show'` (default): drawn like any other slide.
   * - `'skip'`: sequential navigation (`nextSlide`/`prevSlide`, initial load)
   *   jumps over them; absolute indices are unchanged, and an explicit
   *   `goToSlide(i)` to a hidden slide is still honored.
   * - `'dim'`: drawn under a translucent overlay (PowerPoint thumbnail look).
   *
   * Named to match the {@link PptxViewer.hiddenSlideMode} getter and
   * {@link PptxViewer.setHiddenSlideMode} setter.
   */
  hiddenSlideMode?: HiddenSlideMode;
  /**
   * Overrides for the `'dim'` overlay. Merged over the default
   * `{ color: '#ffffff', opacity: 0.6 }`. A `Partial<DimOptions>` so it stays
   * in sync if {@link DimOptions} gains a field.
   */
  hiddenSlideDim?: Partial<DimOptions>;
}

/**
 * Opinionated single-canvas PPTX viewer.
 *
 * Accepts a caller-supplied `<canvas>` element and wraps it in a positioned
 * container for the optional text-selection overlay.  The wrapper is inserted
 * into the canvas's existing parent (reparent), so the canvas stays at its
 * original position in the DOM.
 *
 * For custom layouts (multi-canvas, thumbnails, scroll view) use PptxPresentation directly.
 */
export class PptxViewer {
  private readonly canvas: HTMLCanvasElement;
  private readonly wrapper: HTMLDivElement;
  /** The canvas's DOM position BEFORE the constructor reparented it into
   *  {@link wrapper}, captured so {@link destroy} can return the caller-owned
   *  canvas to exactly where it was. `null` parent = canvas was passed
   *  detached. */
  private readonly _originalParent: Node | null;
  private readonly _originalNextSibling: Node | null;
  /** The canvas's inline `display` before the constructor forced `block`
   *  (empty string if it was unset), restored on {@link destroy}. */
  private readonly _originalDisplay: string;
  private textLayer: HTMLDivElement | null = null;
  private engine: PptxPresentation | null = null;
  private readonly opts: PptxViewerOptions;
  private currentSlide = 0;
  private _hiddenMode: HiddenSlideMode;
  private handle: PresentationHandle | null = null;
  private readonly _mode: 'main' | 'worker';
  /** The canvas's bitmaprenderer context, used only by the static worker-mode
   *  render path. The media-playback path keeps a 2d context (via presentSlide),
   *  so this is obtained only when worker mode renders without media playback. */
  private _bitmapCtx: ImageBitmapRenderingContext | null = null;
  private _warnedNoTextSelection = false;

  constructor(canvas: HTMLCanvasElement, opts: PptxViewerOptions = {}) {
    this.opts = opts;
    this.canvas = canvas;
    this._mode = opts.mode ?? 'main';
    this._hiddenMode = opts.hiddenSlideMode ?? 'show';

    const parent = canvas.parentElement;
    // Capture the canvas's DOM position and inline display BEFORE reparenting so
    // destroy() can put the caller-owned canvas back exactly where it was.
    this._originalParent = parent;
    this._originalNextSibling = canvas.nextSibling;
    this._originalDisplay = canvas.style.display;
    this.wrapper = document.createElement('div');
    // vertical-align:top removes the inline-block baseline descender gap that
    // otherwise lets the host container's background show through below the
    // canvas (~6 px on default font metrics).
    this.wrapper.style.cssText = 'position:relative;display:inline-block;vertical-align:top;';
    // Force `display:block` on the canvas so it does not inherit the inline
    // baseline of an enclosing wrapper, which would otherwise leave a 4–6px
    // descender gap between the canvas bottom and the wrapper bottom — the
    // host container's background would show through that strip.
    if (!canvas.style.display) canvas.style.display = 'block';
    if (parent) parent.insertBefore(this.wrapper, canvas);
    this.wrapper.appendChild(canvas);

    // Static worker-mode rendering paints worker-produced bitmaps via a
    // bitmaprenderer context (grabbed once — a canvas holds one context type for
    // its lifetime). The media-playback path uses presentSlide, which keeps a 2d
    // context, so skip bitmaprenderer there.
    if (this._mode === 'worker' && !opts.enableMediaPlayback) {
      this._bitmapCtx = canvas.getContext('bitmaprenderer');
    }

    if (opts.enableTextSelection) {
      this.textLayer = document.createElement('div');
      this.textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      this.wrapper.appendChild(this.textLayer);
    }
  }

  /**
   * Load a PPTX from URL or ArrayBuffer and render the first slide.
   *
   * Error contract (shared by all three viewers): on failure, if an `onError`
   * callback was provided it is invoked and `load` resolves normally; if not,
   * the error is rethrown so it is never silently swallowed.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    // SC20 atomic swap: retain the previous engine locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // engine's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current engine + its rendered slide intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old engine is freed the moment the new model arrives).
    const previous = this.engine;
    try {
      const engine = await PptxPresentation.load(source, {
        useGoogleFonts: this.opts.useGoogleFonts,
        maxZipEntryBytes: this.opts.maxZipEntryBytes,
        workerTimeoutMs: this.opts.workerTimeoutMs,
        wasmUrl: this.opts.wasmUrl,
        math: this.opts.math,
        mode: this._mode,
      });
      // Discard the stale slide's media handle before swapping engines so its RAF
      // loop / object URLs don't outlive the replaced presentation.
      this.handle?.destroy();
      this.handle = null;
      this.engine = engine;
      previous?.destroy();
      this.currentSlide = this._initialSlide();
      await this.renderCurrentSlide();
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      if (this.opts.onError) {
        this.opts.onError(e);
        return;
      }
      throw e;
    }
  }

  /** Navigate to a specific slide (0-indexed). */
  async goToSlide(index: number): Promise<void> {
    if (!this.engine || this.slideCount === 0) return;
    this.currentSlide = Math.max(0, Math.min(index, this.slideCount - 1));
    await this.renderCurrentSlide();
  }

  async nextSlide(): Promise<void> {
    await this.goToSlide(this._step(1));
  }

  async prevSlide(): Promise<void> {
    await this.goToSlide(this._step(-1));
  }

  /** Next index for sequential nav: skip mode jumps over hidden slides. */
  private _step(dir: 1 | -1): number {
    if (this._hiddenMode === 'skip' && this.engine) {
      return nextVisibleIndex(this.currentSlide, dir, (i) => this.engine!.isHidden(i), this.slideCount);
    }
    return this.currentSlide + dir;
  }

  /** Initial slide for load() / mode switch: skip mode lands on a visible one. */
  private _initialSlide(): number {
    if (this._hiddenMode === 'skip' && this.engine) {
      return resolveVisibleIndex(0, (i) => this.engine!.isHidden(i), this.slideCount);
    }
    return 0;
  }

  /** Resolved `'dim'` overlay (defaults merged with the `hiddenSlideDim` option). */
  private _dim(): DimOptions {
    return {
      color: this.opts.hiddenSlideDim?.color ?? DEFAULT_HIDDEN_DIM.color,
      opacity: this.opts.hiddenSlideDim?.opacity ?? DEFAULT_HIDDEN_DIM.opacity,
    };
  }

  /**
   * Switch the hidden-slide mode at runtime and re-render. Entering `'skip'`
   * while on a hidden slide advances to the nearest visible slide.
   */
  async setHiddenSlideMode(mode: HiddenSlideMode): Promise<void> {
    this._hiddenMode = mode;
    if (mode === 'skip' && this.engine) {
      this.currentSlide = resolveVisibleIndex(
        this.currentSlide,
        (i) => this.engine!.isHidden(i),
        this.slideCount,
      );
    }
    await this.renderCurrentSlide();
  }

  /** The current hidden-slide mode. */
  get hiddenSlideMode(): HiddenSlideMode { return this._hiddenMode; }

  /** Number of non-hidden slides (absolute `slideCount` is unchanged). */
  get visibleSlideCount(): number {
    if (!this.engine) return 0;
    const engine = this.engine;
    return countVisible((i) => engine.isHidden(i), this.slideCount);
  }

  get slideIndex(): number { return this.currentSlide; }
  get slideCount(): number { return this.engine?.slideCount ?? 0; }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5). Passthrough to {@link PptxPresentation.getNotes}:
   * 0-based index, returns `null` when the slide has no notes part, the index
   * is out of range, or nothing is loaded yet.
   */
  getNotes(slideIndex: number): string | null {
    return this.engine?.getNotes(slideIndex) ?? null;
  }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement { return this.canvas; }

  private async renderCurrentSlide(): Promise<void> {
    if (!this.engine) return;
    const dim =
      this._hiddenMode === 'dim' && this.engine.isHidden(this.currentSlide)
        ? this._dim()
        : undefined;
    const targetWidth = this.opts.width ?? (this.canvas.offsetWidth || 960);
    const dpr = this.opts.dpr ?? (window.devicePixelRatio || 1);

    const scale = targetWidth / this.engine.slideWidth;
    const cssHeight = Math.round(this.engine.slideHeight * scale);
    this.canvas.style.width = `${targetWidth}px`;
    this.canvas.style.height = `${cssHeight}px`;

    this.handle?.destroy();
    this.handle = null;

    const isWorker = this._mode === 'worker';
    // In worker mode rendering happens off the main thread, so the onTextRun
    // callback can't fire — the text-selection overlay is unavailable.
    if (isWorker && this.textLayer && !this._warnedNoTextSelection) {
      this._warnedNoTextSelection = true;
      console.warn(
        "[ooxml] text selection is unavailable in mode: 'worker'; the overlay will be empty. Use mode: 'main' for selectable text.",
      );
    }
    const runs: PptxTextRunInfo[] = [];
    const onTextRun = !isWorker && this.textLayer ? (r: PptxTextRunInfo) => runs.push(r) : undefined;

    try {
      if (this.opts.enableMediaPlayback) {
        // presentSlide supports both modes (worker: base off-thread, video
        // overlay composited on the main thread).
        this.handle = await this.engine.presentSlide(this.canvas, this.currentSlide, {
          width: targetWidth,
          dpr,
          dim,
        });
      } else if (isWorker) {
        const bmp = await this.engine.renderSlideToBitmap(this.currentSlide, { width: targetWidth, dpr, dim });
        this.canvas.width = bmp.width;
        this.canvas.height = bmp.height;
        this._bitmapCtx?.transferFromImageBitmap(bmp);
      } else {
        await this.engine.renderSlide(this.canvas, this.currentSlide, { width: targetWidth, dpr, onTextRun, dim });
      }
      this.opts.onSlideChange?.(this.currentSlide, this.slideCount);
    } catch (err) {
      this._reportRenderError(err);
    }

    if (this.textLayer && !isWorker) {
      this._buildTextLayer(this.textLayer, runs, targetWidth, cssHeight);
    }
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: PptxTextRunInfo[], cssWidth: number, cssHeight: number): void {
    buildPptxTextLayer(layer, runs, cssWidth, cssHeight);
  }

  /** PD14 render-error contract: route a render failure to `onError`, or
   *  `console.error` when none is given (never fully silent). Mirrors the scroll
   *  viewers' `_reportRenderError` so all three single-canvas viewers agree. */
  private _reportRenderError(err: unknown): void {
    const e = err instanceof Error ? err : new Error(String(err));
    if (this.opts.onError) this.opts.onError(e);
    else console.error('[ooxml] PptxViewer render failed:', e);
  }

  /**
   * Clean up the viewer and terminate the background worker.
   *
   * The caller-owned `<canvas>` is returned to the DOM position it held before
   * the constructor was called (same parent, same next-sibling) and its inline
   * `display` is restored, so the canvas can be reused — e.g. to construct a new
   * viewer on the same element. If the canvas was passed detached (no parent) it
   * is simply removed from the internal wrapper. Safe to call more than once.
   */
  destroy(): void {
    this.handle?.destroy();
    this.handle = null;
    this.engine?.destroy();
    // Return the caller-owned canvas to its original DOM slot before discarding
    // the wrapper. insertBefore still works if the original parent was itself
    // detached; when there was no original parent the canvas is left detached
    // (just pulled out of the wrapper). The recorded next-sibling may have been
    // removed or moved by the caller since construction — insertBefore throws
    // NotFoundError for a reference that is no longer a child of the parent, so
    // fall back to appending at the end in that case.
    if (this._originalParent) {
      const ref =
        this._originalNextSibling && this._originalNextSibling.parentNode === this._originalParent
          ? this._originalNextSibling
          : null;
      this._originalParent.insertBefore(this.canvas, ref);
    } else if (this.canvas.parentNode) {
      this.canvas.parentNode.removeChild(this.canvas);
    }
    this.canvas.style.display = this._originalDisplay;
    this.wrapper.remove();
  }
}
