import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { RenderPageOptions } from './types';
import type { DocxTextRunInfo } from './renderer';
import { buildDocxTextLayer } from './text-layer';

export interface DocxViewerOptions extends RenderPageOptions, LoadOptions {
  container?: HTMLElement;
  /**
   * When true, adds a transparent text overlay div over the canvas so the
   * browser's native text selection works on document content.
   */
  enableTextSelection?: boolean;
  /** Called when a page finishes rendering. */
  onPageChange?: (index: number, total: number) => void;
  /** Called on parse or render errors. */
  onError?: (err: Error) => void;
}

export class DocxViewer {
  private _doc: DocxDocument | null = null;
  private _currentPage = 0;
  private _canvas: HTMLCanvasElement;
  private _wrapper: HTMLDivElement;
  private _textLayer: HTMLDivElement | null = null;
  private _opts: DocxViewerOptions;
  private readonly _mode: 'main' | 'worker';
  /** The canvas's bitmaprenderer context, used only in worker mode (a canvas
   *  holds one context type for its lifetime; the main-mode 2d render path is
   *  never used on the same canvas). */
  private _bitmapCtx: ImageBitmapRenderingContext | null = null;
  private _warnedNoTextSelection = false;

  constructor(canvas: HTMLCanvasElement, opts: DocxViewerOptions = {}) {
    this._canvas = canvas;
    this._opts = opts;
    this._mode = opts.mode ?? 'main';

    // Wrap canvas in a positioned container for the optional text layer overlay
    const parent = canvas.parentElement;
    this._wrapper = document.createElement('div');
    // vertical-align:top removes the inline-block baseline descender gap that
    // otherwise lets the host container's background show through below the
    // canvas (~6 px on default font metrics).
    this._wrapper.style.cssText = 'position:relative;display:inline-block;vertical-align:top;';
    // Force `display:block` on the canvas so it does not inherit the inline
    // baseline of the wrapper, which would otherwise leave a 4–6px descender
    // gap between the canvas bottom and the wrapper bottom — the host
    // container's background would show through that strip.
    if (!canvas.style.display) canvas.style.display = 'block';
    if (parent) {
      parent.insertBefore(this._wrapper, canvas);
    }
    this._wrapper.appendChild(canvas);

    // Worker mode paints worker-produced bitmaps via a bitmaprenderer context,
    // grabbed once (a canvas holds one context type for its lifetime).
    if (this._mode === 'worker') {
      this._bitmapCtx = canvas.getContext('bitmaprenderer');
    }

    if (opts.enableTextSelection) {
      this._textLayer = document.createElement('div');
      this._textLayer.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;' +
        'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';
      this._wrapper.appendChild(this._textLayer);
    }
  }

  /**
   * Load a DOCX from URL or ArrayBuffer and render the first page.
   *
   * Error contract (shared by all three viewers): on failure, if an `onError`
   * callback was provided it is invoked and `load` resolves normally; if not,
   * the error is rethrown so it is never silently swallowed.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    try {
      this._doc = await DocxDocument.load(source, {
        useGoogleFonts: this._opts.useGoogleFonts,
        maxZipEntryBytes: this._opts.maxZipEntryBytes,
        math: this._opts.math,
        mode: this._mode,
      });
      this._currentPage = 0;
      await this._render();
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

  get currentPage(): number {
    return this._currentPage;
  }

  /** The underlying <canvas> element. */
  get canvasElement(): HTMLCanvasElement {
    return this._canvas;
  }

  async goToPage(index: number): Promise<void> {
    if (!this._doc) return;
    const clamped = Math.max(0, Math.min(index, this.pageCount - 1));
    this._currentPage = clamped;
    await this._render();
  }

  async nextPage(): Promise<void> { await this.goToPage(this._currentPage + 1); }
  async prevPage(): Promise<void> { await this.goToPage(this._currentPage - 1); }

  /** Terminate the parser worker and release resources. */
  destroy(): void {
    this._doc?.destroy();
    this._doc = null;
    this._wrapper.remove();
  }

  private async _render(): Promise<void> {
    if (!this._doc) return;
    const isWorker = this._mode === 'worker';
    // In worker mode rendering happens off the main thread, so the onTextRun
    // callback can't fire — the text-selection overlay is unavailable.
    if (isWorker && this._textLayer && !this._warnedNoTextSelection) {
      this._warnedNoTextSelection = true;
      console.warn(
        "[ooxml] text selection is unavailable in mode: 'worker'; the overlay will be empty. Use mode: 'main' for selectable text.",
      );
    }
    if (isWorker) {
      const dpr = this._opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
      // Only serializable render options may cross to the worker — spreading the
      // full viewer opts would postMessage non-cloneable values (the math
      // engine, callbacks, container element) and throw a DataCloneError.
      const bmp = await this._doc.renderPageToBitmap(this._currentPage, {
        width: this._opts.width,
        dpr: this._opts.dpr,
        showTrackChanges: this._opts.showTrackChanges,
      });
      this._canvas.width = bmp.width;
      this._canvas.height = bmp.height;
      // The bitmap is sized in device px; mirror the main renderer by setting
      // the CSS size to the logical (÷dpr) dimensions so it isn't 2× on HiDPI.
      this._canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
      this._canvas.style.height = `${Math.round(bmp.height / dpr)}px`;
      this._bitmapCtx?.transferFromImageBitmap(bmp);
    } else {
      const runs: DocxTextRunInfo[] = [];
      const onTextRun = this._textLayer ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
      await this._doc.renderPage(this._canvas, this._currentPage, { ...this._opts, onTextRun });
      if (this._textLayer) {
        this._buildTextLayer(this._textLayer, runs);
      }
    }
    this._opts.onPageChange?.(this._currentPage, this.pageCount);
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: DocxTextRunInfo[]): void {
    buildDocxTextLayer(
      layer,
      runs,
      this._canvas.style.width || this._canvas.width + 'px',
      this._canvas.style.height || this._canvas.height + 'px',
    );
  }
}
