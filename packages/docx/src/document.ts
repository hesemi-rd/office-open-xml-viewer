import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import type { PaginatedBodyElement, DocxDocumentModel, RenderPageOptions, WorkerRequest, WorkerResponse, DocComment, DocNote } from './types';
import { renderDocumentToCanvas, documentHasMath, prepareMathRuns, paginateDocument } from './renderer';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import type {
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
  WireRenderPageOptions,
} from './worker-protocol';

/** Options for {@link DocxDocument.load}. Extends the shared load-options type
 *  from `@silurus/ooxml-core` (`useGoogleFonts`, `maxZipEntryBytes`) with the
 *  opt-in math engine. */
export interface LoadOptions extends CoreLoadOptions {
  /**
   * Opt-in OMML equation engine. Import it from the separate `@silurus/ooxml/math`
   * entry and pass it in: `import { math } from '@silurus/ooxml/math'`. When
   * omitted, equations are skipped and the ~3 MB engine never enters the bundle.
   */
  math?: MathRenderer;
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse, paginate AND render inside the worker; use
   * {@link DocxDocument.renderPageToBitmap} and paint the returned ImageBitmap
   * via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas. The math
   * engine is unavailable in this mode (equations are skipped).
   */
  mode?: 'main' | 'worker';
}

export class DocxDocument {
  private _document: DocxDocumentModel | null = null;
  private _meta: DocumentMeta | null = null;
  private _pages: PaginatedBodyElement[][] | null = null;
  private _mode: 'main' | 'worker' = 'main';
  private _worker: Worker;
  private _bridge: WorkerBridge<WorkerResponse | RenderWorkerResponse>;
  private _imageCache = new Map<string, Promise<Blob>>();

  private constructor(worker: Worker, mode: 'main' | 'worker') {
    this._worker = worker;
    this._mode = mode;
    this._bridge = new WorkerBridge<WorkerResponse | RenderWorkerResponse>(this._worker, {
      correlate: (res) => res.id,
      toError: (res) => (res.type === 'error' ? res.message : undefined),
    });
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._bridge.post({ type: 'init', wasmUrl } satisfies WorkerRequest);
  }

  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<DocxDocument> {
    const mode = opts.mode ?? 'main';
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const doc = new DocxDocument(worker, mode);
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    if (opts.math && mode === 'worker') {
      console.warn(
        "[ooxml] the math engine is unavailable in mode: 'worker'; equations will be skipped. Use mode: 'main' for documents with equations.",
      );
    }
    // In worker mode the worker preloads fonts before paginating (pagination
    // measures text), so the flag is forwarded; in main mode fonts are loaded
    // here after parse, before the lazy first pagination.
    await doc._parse(
      buffer,
      opts.maxZipEntryBytes,
      mode === 'worker' ? !!opts.useGoogleFonts : false,
    );
    if (mode === 'main' && opts.useGoogleFonts && doc._document) {
      await preloadGoogleFonts(
        docxFontPreloadNames(doc._document),
        DOCX_GOOGLE_FONTS,
      );
    }
    // Equations are converted + rasterized before pagination (which reads their
    // extents synchronously). Requires the opt-in `math` engine; without it,
    // equations are skipped (and the engine asset is never bundled). Math is
    // main-mode only (the engine needs a DOM, absent in workers).
    if (mode === 'main' && opts.math && doc._document && documentHasMath(doc._document.body)) {
      await prepareMathRuns(doc._document.body, opts.math);
    }
    return doc;
  }

  private async _parse(
    buffer: ArrayBuffer,
    maxZipEntryBytes?: number,
    useGoogleFonts = false,
  ): Promise<void> {
    const res = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ type: 'parse', id, data: buffer, maxZipEntryBytes, useGoogleFonts } satisfies RenderWorkerRequest)
          : ({ type: 'parse', id, data: buffer, maxZipEntryBytes } satisfies WorkerRequest),
      [buffer],
    );
    if (this._mode === 'worker') {
      this._meta = (res as Extract<RenderWorkerResponse, { type: 'parsedMeta' }>).meta;
    } else {
      this._document = (res as Extract<WorkerResponse, { type: 'parsed' }>).document;
    }
  }

  destroy(): void {
    this._bridge.terminate();
    this._document = null;
    this._meta = null;
    this._pages = null;
    this._imageCache.clear();
  }

  /**
   * Extract raw bytes for an embedded image by zip path (e.g.
   * `word/media/image1.png`), wrapped in a Blob of the given MIME type. Routes
   * through the persistent worker via the `extractImage` message (twin of
   * pptx's `getImage`/`getMedia`); results are cached by path for the lifetime
   * of this instance. The renderer's `fetchImage` option points here so images
   * are decoded lazily rather than inlined as base64 at parse time.
   */
  async getImage(imagePath: string, mimeType: string): Promise<Blob> {
    const hit = this._imageCache.get(imagePath);
    if (hit) return hit;
    const p = this._bridge
      .request((id) => ({ type: 'extractImage', id, path: imagePath }) satisfies WorkerRequest)
      .then((res) => {
        const bytes = (res as Extract<WorkerResponse, { type: 'imageExtracted' }>).bytes;
        return new Blob([bytes], { type: mimeType });
      });
    this._imageCache.set(imagePath, p);
    return p;
  }

  get pageCount(): number {
    if (this._meta) return this._meta.pageCount;
    if (!this._document) return 0;
    return this._getPages().length;
  }

  /**
   * The raw parsed document model. Available only in `mode: 'main'`; in
   * `mode: 'worker'` the model stays in the worker and this throws.
   */
  get document(): DocxDocumentModel {
    if (this._meta && !this._document) {
      throw new Error(
        "the raw document model stays in the worker in mode: 'worker'; use mode: 'main' if you need direct model access",
      );
    }
    if (!this._document) throw new Error('Document not loaded');
    return this._document;
  }

  /**
   * ECMA-376 §17.13.4 — the document's comments (`word/comments.xml`), each with
   * id / author / initials / date / plain-text body. Comments are a data-only
   * API: they are NOT drawn on the page (Word renders them in a margin pane /
   * balloons, which this viewer does not reproduce). Use this to build a review
   * panel, export an annotation list, etc. Returns `[]` when the document has no
   * comments part. The same data is also reachable via `document.comments`.
   */
  get comments(): DocComment[] {
    return this._meta?.comments ?? this._document?.comments ?? [];
  }

  /**
   * ECMA-376 §17.11.10 — the document's footnotes (`word/footnotes.xml`),
   * excluding the reserved separator entries. Each note carries its `id` and
   * block-level `content`; use {@link noteText} for the plain-text body. These
   * ARE drawn at the bottom of the page that holds their reference; this getter
   * additionally exposes them as data. Returns `[]` when absent.
   */
  get footnotes(): DocNote[] {
    return this._meta?.footnotes ?? this._document?.footnotes ?? [];
  }

  /**
   * ECMA-376 §17.11.4 — the document's endnotes (`word/endnotes.xml`). Same
   * shape as {@link footnotes}; rendered at the end of the document. Returns
   * `[]` when absent.
   */
  get endnotes(): DocNote[] {
    return this._meta?.endnotes ?? this._document?.endnotes ?? [];
  }

  private _getPages(): PaginatedBodyElement[][] {
    if (this._pages) return this._pages;
    if (!this._document) return [];
    this._pages = paginateDocument(this._document);
    return this._pages;
  }

  renderPage(
    target: HTMLCanvasElement | OffscreenCanvas,
    pageIndex: number,
    opts: RenderPageOptions = {},
  ): Promise<void> {
    if (this._mode === 'worker') {
      throw new Error(
        "renderPage(canvas) is unavailable in mode: 'worker'; use renderPageToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this._document) throw new Error('Document not loaded');
    const pages = this._getPages();
    return renderDocumentToCanvas(this._document, target, pageIndex, {
      ...opts,
      totalPages: pages.length,
      prebuiltPages: pages,
      // Lazy image bytes: the renderer fetches each embedded blip on demand by
      // zip path (decoded only when drawn) instead of reading inlined base64.
      fetchImage: (path, mime) => this.getImage(path, mime),
    });
  }

  /**
   * Render a page and return it as an ImageBitmap. Works in both modes; in
   * worker mode the render runs entirely off the main thread. Paint with:
   * `canvas.getContext('bitmaprenderer').transferFromImageBitmap(bitmap)`.
   *
   * The returned ImageBitmap is owned by the caller: pass it to
   * `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`
   * when done, or its backing memory is held until GC.
   */
  async renderPageToBitmap(pageIndex: number, opts: WireRenderPageOptions = {}): Promise<ImageBitmap> {
    const wireOpts = { ...opts, dpr: opts.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= this.pageCount) {
        throw new Error(`Page index ${pageIndex} out of range (count: ${this.pageCount})`);
      }
      const res = await this._bridge.request(
        (id) => ({ type: 'renderPage', id, pageIndex, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { type: 'pageRendered' }>).bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderPage(off, pageIndex, wireOpts);
    return off.transferToImageBitmap();
  }
}
