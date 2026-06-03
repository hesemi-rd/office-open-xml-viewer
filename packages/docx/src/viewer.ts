import { DocxDocument } from './document';
import type { LoadOptions } from './document';
import type { RenderPageOptions } from './types';
import type { DocxTextRunInfo } from './renderer';

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

  constructor(canvas: HTMLCanvasElement, opts: DocxViewerOptions = {}) {
    this._canvas = canvas;
    this._opts = opts;

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
    const runs: DocxTextRunInfo[] = [];
    const onTextRun = this._textLayer ? (r: DocxTextRunInfo) => runs.push(r) : undefined;
    await this._doc.renderPage(this._canvas, this._currentPage, { ...this._opts, onTextRun });
    if (this._textLayer) {
      this._buildTextLayer(this._textLayer, runs);
    }
    this._opts.onPageChange?.(this._currentPage, this.pageCount);
  }

  private _buildTextLayer(layer: HTMLDivElement, runs: DocxTextRunInfo[]): void {
    layer.innerHTML = '';
    layer.style.width = `${this._canvas.style.width || this._canvas.width + 'px'}`;
    layer.style.height = `${this._canvas.style.height || this._canvas.height + 'px'}`;

    for (const run of runs) {
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
        `left:${run.x}px;top:${run.y}px;` +
        `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
        `white-space:pre;color:transparent;cursor:text;pointer-events:all;`;
      layer.appendChild(span);
    }
  }
}
