import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  unregisterEmbeddedFonts,
  WorkerBridge,
  defaultDpr,
  dropSvgImageCache,
  dropBitmapCacheByPath,
  resolveOoxmlContainer,
  toArrayBuffer,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import type { PaginatedBodyElement, DocxDocumentModel, RenderPageOptions, WorkerRequest, WorkerResponse, DocComment, DocNote } from './types';
import { renderDocumentToCanvas, documentHasMath, prepareMathRuns, paginateDocument, dropColorReplacedCache, physicalPageSizeForPage, type DocxTextRunInfo } from './renderer';
import { buildBookmarkPageMap } from './bookmark-nav';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
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

/** IX6 — options for {@link DocxDocument.renderPageToBitmap}: the serializable
 *  render knobs plus an OPTIONAL `onTextRun`. The callback stays main-thread (it
 *  never crosses the wire); in worker mode the proxy invokes it with the runs
 *  the worker shipped back beside the bitmap, so a caller gets the selection /
 *  find geometry on the same path in both modes. */
export type RenderPageToBitmapOptions = WireRenderPageOptions & {
  onTextRun?: (run: DocxTextRunInfo) => void;
};

export class DocxDocument {
  private _document: DocxDocumentModel | null = null;
  private _meta: DocumentMeta | null = null;
  private _pages: PaginatedBodyElement[][] | null = null;
  /** Lazily-built `bookmarkName → 0-based page index` map for internal hyperlink
   *  anchors (IX-nav). Built on first {@link getBookmarkPage} from the paginated
   *  pages (main) or the worker meta's `bookmarkPages` (worker). Nulled by
   *  {@link destroy} so a reused reference never serves a stale document. */
  private _bookmarkPages: Map<string, number> | null = null;
  private _mode: 'main' | 'worker' = 'main';
  private _worker: Worker;
  private _bridge: WorkerBridge<WorkerResponse | RenderWorkerResponse>;
  private _imageCache = new Map<string, Promise<Blob>>();
  /** Embedded `FontFace` objects this document registered into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with
   *  its own FontFaceSet). Released in {@link destroy} so they do not leak into
   *  the shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in
   *  core, so a font shared with another open document survives until both go). */
  private _embeddedFontFaces: FontFace[] = [];
  /** Google-Fonts `FontFace` objects this document preloaded into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with its
   *  own FontFaceSet). Released in {@link destroy} so they do not leak into the
   *  shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in core,
   *  so a web font shared with another open document survives until both go). */
  private _googleFontFaces: FontFace[] = [];
  /** One stable closure per instance: core's path-keyed SVG cache namespaces on
   *  this identity, so two open documents never swap a shared zip path (e.g.
   *  word/media/image1.svg). Reusing one reference also lets the SVG cache hit
   *  across page renders. */
  private readonly _fetchImage = (path: string, mime: string): Promise<Blob> =>
    this.getImage(path, mime);

  private constructor(worker: Worker, mode: 'main' | 'worker', wasmUrlOverride?: string | URL) {
    this._worker = worker;
    this._mode = mode;
    this._bridge = new WorkerBridge<WorkerResponse | RenderWorkerResponse>(this._worker, {
      correlate: (res) => res.id,
      toError: (res) => (res.type === 'error' ? res.message : undefined),
    });
    // Default: the parser WASM emitted next to this bundle, resolved relative to
    // the document URL. `wasmUrl` overrides it (CDN / self-hosted copy); a
    // relative override is still resolved against `location.href`.
    const wasmUrl = new URL(wasmUrlOverride ?? wasmAssetUrl, location.href).href;
    this._bridge.post({ type: 'init', wasmUrl } satisfies WorkerRequest);
  }

  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<DocxDocument> {
    const mode = opts.mode ?? 'main';
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    // Resolve the container on the main thread — before spinning up the worker.
    // A normal ZIP passes through unchanged; an Agile-encrypted CFB is decrypted
    // in place when `opts.password` is supplied ([MS-OFFCRYPTO]); a
    // password-protected file without a password, or a legacy-binary / unknown
    // CFB, becomes a typed OoxmlError (whose `instanceof` would not survive the
    // worker boundary, hence the main-thread check).
    buffer = toArrayBuffer(await resolveOoxmlContainer(buffer, opts.password));
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const doc = new DocxDocument(worker, mode, opts.wasmUrl);
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
      opts.workerTimeoutMs,
    );
    if (mode === 'main' && opts.useGoogleFonts && doc._document) {
      doc._googleFontFaces = await preloadGoogleFonts(
        docxFontPreloadNames(doc._document),
        DOCX_GOOGLE_FONTS,
      );
    }
    // ECMA-376 §17.8.1 / §17.8.3 — register the document's embedded fonts (via
    // the worker's zip-entry extraction) before the lazy first pagination, so
    // text measures/draws with the authored typeface. Worker mode does this
    // inside the worker (before it paginates); here it runs on the main thread.
    if (mode === 'main' && doc._document?.embeddedFonts?.length) {
      doc._embeddedFontFaces = await loadEmbeddedFonts(doc._document, (p) => doc.getFontBytes(p));
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
    timeoutMs?: number,
  ): Promise<void> {
    const res = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ type: 'parse', id, data: buffer, maxZipEntryBytes, useGoogleFonts } satisfies RenderWorkerRequest)
          : ({ type: 'parse', id, data: buffer, maxZipEntryBytes } satisfies WorkerRequest),
      [buffer],
      { timeoutMs },
    );
    if (this._mode === 'worker') {
      this._meta = (res as Extract<RenderWorkerResponse, { type: 'parsedMeta' }>).meta;
    } else {
      // The model arrives as transferred UTF-8 JSON bytes; decode + parse once
      // here (the only serialization on the parse-mode path).
      const { documentJson } = res as Extract<WorkerResponse, { type: 'parsed' }>;
      this._document = JSON.parse(
        new TextDecoder().decode(new Uint8Array(documentJson)),
      ) as DocxDocumentModel;
    }
  }

  destroy(): void {
    this._bridge.terminate();
    this._document = null;
    this._meta = null;
    this._pages = null;
    this._bookmarkPages = null;
    this._imageCache.clear();
    // Release the embedded fonts this document added to the shared FontFaceSet
    // (main mode). Refcounted in core: a font also used by another open document
    // stays until that one is destroyed too. Without this, every opened document
    // left its FontFace objects in `document.fonts` forever (SPA memory leak).
    if (this._embeddedFontFaces.length > 0) {
      unregisterEmbeddedFonts(this._embeddedFontFaces);
      this._embeddedFontFaces = [];
    }
    // Release the Google-Fonts substitutes this document preloaded into the
    // shared FontFaceSet (main mode). Same refcount contract as the embedded
    // fonts: a web font also used by another open document stays until that one
    // is destroyed too. Without this, every opened document left its Google
    // FontFace objects in `document.fonts` forever (SPA memory leak).
    if (this._googleFontFaces.length > 0) {
      unloadGoogleFonts(this._googleFontFaces);
      this._googleFontFaces = [];
    }
    // Release this document's three per-fetchImage image caches: the decoded
    // base raster bitmaps (GPU-backed, shared with pptx/xlsx), the a:clrChange
    // color-replaced bitmaps (docx-only second layer), and the decoded-SVG
    // object URLs. All are keyed by `_fetchImage`, so a discarded document frees
    // its GPU/URL handles promptly rather than waiting for GC.
    dropBitmapCacheByPath(this._fetchImage);
    dropColorReplacedCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
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

  /**
   * Extract raw bytes for an embedded font part by zip path (e.g.
   * `word/fonts/font1.odttf`). Routes through the SAME persistent-worker
   * `extractImage` message as {@link getImage} — `DocxArchive.extract_image`
   * reads ANY zip entry, not just media — returning the raw (still obfuscated)
   * `.odttf` bytes rather than a Blob. Consumed by {@link loadEmbeddedFonts},
   * which de-obfuscates (ECMA-376 §17.8.1) and registers each as a FontFace.
   */
  async getFontBytes(partPath: string): Promise<Uint8Array> {
    const res = await this._bridge.request(
      (id) => ({ type: 'extractImage', id, path: partPath }) satisfies WorkerRequest,
    );
    const bytes = (res as Extract<WorkerResponse, { type: 'imageExtracted' }>).bytes;
    return new Uint8Array(bytes);
  }

  /**
   * Project the document to GitHub-flavoured markdown: headings (from
   * `<w:outlineLvl>`), bullet / numbered lists, tables (with vMerge
   * continuation), and rich-text formatting (bold / italic / strikethrough /
   * hyperlink), with footnotes / endnotes / comments collated at the end.
   * Positioning, section properties, fonts, and drawing shapes are discarded —
   * the projection is meant for AI ingestion and full-text search, not layout.
   *
   * Runs entirely in the worker off the archive opened at {@link load} (no
   * re-copy of the file, no re-parse of the model on the main thread), so it
   * works in BOTH `mode: 'main'` and `mode: 'worker'`.
   *
   * @example
   * const doc = await DocxDocument.load(buffer);
   * const md = await doc.toMarkdown();
   */
  async toMarkdown(): Promise<string> {
    const res = await this._bridge.request(
      (id) => ({ type: 'toMarkdown', id }) satisfies WorkerRequest,
    );
    return (res as Extract<WorkerResponse, { type: 'markdownRendered' }>).markdown;
  }

  get pageCount(): number {
    if (this._meta) return this._meta.pageCount;
    if (!this._document) return 0;
    return this._getPages().length;
  }

  /** The render mode this engine was loaded with ('main' | 'worker'). A fact for
   *  integrators and the scroll viewer: an injected engine's mode decides whether
   *  pages render via renderPage (main) or renderPageToBitmap (worker) — no
   *  probing (design §11: no silent mis-pathing). */
  get mode(): 'main' | 'worker' {
    return this._mode;
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

  /** Lazily build (and cache) the `bookmarkName → page index` map from either
   *  the worker meta (worker mode) or the paginated pages (main mode). */
  private _getBookmarkPages(): Map<string, number> {
    if (this._bookmarkPages) return this._bookmarkPages;
    this._bookmarkPages = this._meta
      ? new Map(this._meta.bookmarkPages)
      : buildBookmarkPageMap(this._getPages());
    return this._bookmarkPages;
  }

  /**
   * ECMA-376 §17.13.6.2 / §17.16.23 — resolve a bookmark name (a
   * `<w:hyperlink w:anchor>` internal-link target) to the 0-based index of the
   * page its `<w:bookmarkStart w:name>` destination falls on, or `undefined`
   * when the document has no bookmark of that name. When a bookmark's paragraph
   * spans a page break, the page where it *begins* is returned.
   *
   * This is the map an internal-hyperlink click resolves against: a viewer's
   * `onHyperlinkClick` default (or an integrator) turns the anchor into a page
   * and calls {@link DocxViewer.goToPage} (or scrolls the scroll viewer to it).
   * Works in BOTH `main` and `worker` mode (the map rides along in the worker
   * meta, built from the same paginated pages as `pageSizes`).
   */
  getBookmarkPage(bookmarkName: string): number | undefined {
    return this._getBookmarkPages().get(bookmarkName);
  }

  /**
   * ECMA-376 §17.6.13 / §17.6.11 — the page size (pt) of page `pageIndex`, per
   * section (a mixed portrait/landscape document returns different sizes per page).
   * Available in BOTH modes: worker mode reads the worker-built `pageSizes` meta;
   * main mode reads the paginated pages' stamped geometry. Returns the body-level
   * section size for an out-of-range index (clamped) or a page with no stamped
   * geometry. `{ 0, 0 }` means "not loaded" (before `load()` resolves or after
   * `destroy()`). Returns a fresh object per call — safe to mutate.
   * The recommended way to ask "how big is page i?" for layout.
   */
  pageSize(pageIndex: number): { widthPt: number; heightPt: number } {
    if (this._meta) {
      const sizes = this._meta.pageSizes;
      const clamped = Math.max(0, Math.min(pageIndex, sizes.length - 1));
      const s = sizes[clamped];
      // Copy — never alias the meta's stored object (a caller mutating the
      // return value must not corrupt subsequent reads; main mode below already
      // builds a fresh object per call).
      return s ? { widthPt: s.widthPt, heightPt: s.heightPt } : { widthPt: 0, heightPt: 0 };
    }
    if (!this._document) return { widthPt: 0, heightPt: 0 };
    const pages = this._getPages();
    const clamped = Math.max(0, Math.min(pageIndex, pages.length - 1));
    // The stamped `sectionGeom` is in LOGICAL dims for a vertical section
    // (pagination runs on the swapped frame); `physicalPageSizeForPage` — the
    // resolver shared with the render worker's `pageSizes` meta — un-swaps it
    // by the PAGE's OWN stamped direction (§17.6.20 is per-section, issue
    // #1000), so this reports the visible PHYSICAL page box (identity for
    // horizontal pages).
    return physicalPageSizeForPage(pages, clamped, this._document.section);
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
      fetchImage: this._fetchImage,
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
   *
   * IX6 — an optional `onTextRun` in `opts` receives the page's text-run
   * geometry (the same stream `renderPage` emits in main mode), so a caller can
   * build the selection / find overlay from a worker-rendered page on the SAME
   * code path as main mode. In worker mode the runs ride back beside the bitmap
   * (one round-trip, no second render).
   */
  async renderPageToBitmap(
    pageIndex: number,
    opts: RenderPageToBitmapOptions = {},
  ): Promise<ImageBitmap> {
    const { onTextRun, ...wire } = opts;
    const wireOpts: WireRenderPageOptions = { ...wire, dpr: wire.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= this.pageCount) {
        throw new Error(`Page index ${pageIndex} out of range (count: ${this.pageCount})`);
      }
      const res = await this._bridge.request(
        (id) => ({ type: 'renderPage', id, pageIndex, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      const rendered = res as Extract<RenderWorkerResponse, { type: 'pageRendered' }>;
      if (onTextRun) for (const r of rendered.runs) onTextRun(r);
      return rendered.bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderPage(off, pageIndex, { ...wireOpts, onTextRun });
    return off.transferToImageBitmap();
  }

  /**
   * IX6 — collect a page's text-run geometry (`DocxTextRunInfo[]`) without
   * painting a visible canvas. Works in BOTH modes: worker mode renders the page
   * off-thread and ships only the runs (no bitmap transfer); main mode renders
   * to a throwaway offscreen canvas. Used by the find controller to scan every
   * page for matches. The geometry is identical to a `renderPage` of the same
   * page at the same width/dpr.
   */
  async collectPageRuns(
    pageIndex: number,
    opts: WireRenderPageOptions = {},
  ): Promise<DocxTextRunInfo[]> {
    const wireOpts: WireRenderPageOptions = { ...opts, dpr: opts.dpr ?? defaultDpr() };
    if (this._mode === 'worker') {
      if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= this.pageCount) {
        throw new Error(`Page index ${pageIndex} out of range (count: ${this.pageCount})`);
      }
      const res = await this._bridge.request(
        (id) => ({ type: 'collectRuns', id, pageIndex, opts: wireOpts }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { type: 'runsCollected' }>).runs;
    }
    const runs: DocxTextRunInfo[] = [];
    const off =
      typeof OffscreenCanvas !== 'undefined'
        ? new OffscreenCanvas(1, 1)
        : (globalThis.document?.createElement('canvas') as HTMLCanvasElement);
    await this.renderPage(off, pageIndex, { ...wireOpts, onTextRun: (r) => runs.push(r) });
    return runs;
  }
}
