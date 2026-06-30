import type { RenderOptions, PptxTextRunInfo } from './renderer';
import { PptxPresentation, type LoadOptions } from './presentation';
import type { PresentationHandle } from './presentation-handle';
import { nextVisibleIndex, resolveVisibleIndex } from './hidden';
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
    try {
      this.engine = await PptxPresentation.load(source, {
        useGoogleFonts: this.opts.useGoogleFonts,
        maxZipEntryBytes: this.opts.maxZipEntryBytes,
        math: this.opts.math,
        mode: this._mode,
      });
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
    let n = 0;
    for (let i = 0; i < this.slideCount; i++) if (!this.engine.isHidden(i)) n++;
    return n;
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
      this.opts.onError?.(err instanceof Error ? err : new Error(String(err)));
    }

    if (this.textLayer && !isWorker) {
      this._buildTextLayer(this.textLayer, runs, targetWidth, cssHeight);
    }
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: PptxTextRunInfo[], cssWidth: number, cssHeight: number): void {
    layer.innerHTML = '';
    layer.style.width = `${cssWidth}px`;
    layer.style.height = `${cssHeight}px`;

    // Group runs by shape (same shapeX/shapeY/rotation)
    type ShapeKey = string;
    const shapeMap = new Map<ShapeKey, { div: HTMLDivElement; x: number; y: number; w: number; h: number; rot: number }>();

    for (const run of runs) {
      const totalRot = run.rotation + (run.textBodyRotation ?? 0);
      const key = `${run.shapeX},${run.shapeY},${run.shapeW},${run.shapeH},${totalRot}`;
      if (!shapeMap.has(key)) {
        const div = document.createElement('div');
        div.style.cssText =
          `position:absolute;` +
          `left:${run.shapeX}px;top:${run.shapeY}px;` +
          `width:${run.shapeW}px;height:${run.shapeH}px;` +
          `pointer-events:all;overflow:hidden;`;
        if (totalRot !== 0) {
          div.style.transformOrigin = 'center center';
          div.style.transform = `rotate(${totalRot}deg)`;
        }
        shapeMap.set(key, { div, x: run.shapeX, y: run.shapeY, w: run.shapeW, h: run.shapeH, rot: totalRot });
        layer.appendChild(div);
      }

      const shape = shapeMap.get(key)!;
      const span = document.createElement('span');
      span.textContent = run.text;
      // The `font` shorthand must precede `line-height` because the shorthand
      // resets `line-height` to `normal`. Reset `letter-spacing` so a parent
      // CSS rule cannot drift the trailing edge of the selection. Kerning /
      // ligatures are left at the browser default ('auto') because canvas
      // `measureText` / `fillText` also apply them by default — forcing them
      // off here would make the span wider than the drawn text.
      span.style.cssText =
        `position:absolute;` +
        `left:${run.inShapeX}px;top:${run.inShapeY}px;` +
        `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
        `white-space:pre;color:transparent;cursor:text;`;
      shape.div.appendChild(span);
    }
  }

  /** Clean up the viewer and terminate the background worker. */
  destroy(): void {
    this.handle?.destroy();
    this.handle = null;
    this.engine?.destroy();
    this.wrapper.remove();
  }
}
