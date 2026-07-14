/**
 * Render-capable worker entry: parse → font preload → paginate, all
 * worker-side; renders pages into an OffscreenCanvas and replies with
 * transferable ImageBitmaps. Used by DocxDocument.load(src, { mode: 'worker' });
 * the slim parse-only worker.ts stays untouched so main-mode users pay no
 * bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 */
import init, { DocxArchive, reinit } from './wasm/docx_parser.js';
import {
  decodeDataUrl,
  preloadGoogleFonts,
  unloadLocalFontMetrics,
  WasmParserHost,
  dropBitmapCacheByPath,
  dropSvgImageCache,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel, PaginatedBodyElement } from './types';
import { paginateDocument, renderDocumentToCanvas, createLayoutServices, physicalPageSizeForPage, dropColorReplacedCache, type DocxTextRunInfo } from './renderer';
import { buildBookmarkPageMap } from './bookmark-nav';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
import { loadDocxLocalFontMetrics } from './local-font-metrics';
import type { LayoutServices } from './layout/types.js';
import type { DocumentLayout } from './layout/types.js';
import { layoutParseErrorPage } from './layout/error-page.js';
import { deepFreezeDocumentLayout } from './layout/invariants.js';
import type { RenderWorkerRequest, RenderWorkerResponse, DocumentMeta } from './worker-protocol';
import { normalizeInternalDocumentModel } from './parser-model.js';

// RB6: self-poison + auto-respawn. A trap during parse (or an in-worker image /
// embedded-font read) recycles the instance so the next document renders on
// clean linear memory. The host owns the `DocxArchive` handle (`host.archive`).
const host = new WasmParserHost<DocxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let doc: {
  model: DocxDocumentModel;
  layoutServices: LayoutServices;
  defaultCurrentDateMs: number;
  retainedErrorLayout: DocumentLayout | null;
} | null = null;
let pages: PaginatedBodyElement[][] | null = null;
let localMetricFontFaces: FontFace[] = [];
const imageCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** In-worker image-byte loader (twin of pptx's render-worker `getImage`). The
 *  renderer's `fetchImage` routes here in worker mode, so image bytes are
 *  decoded straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    const loaded = host.archive;
    if (!loaded) throw new Error('No docx loaded');
    const bytes = host.run(() => loaded.extract_image(path));
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageCache.set(path, p);
  return p;
}

self.onmessage = async (e: MessageEvent<RenderWorkerRequest>) => {
  const req = e.data;
  if (req.type === 'init') {
    host.setWasmUrl(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }
  const id = req.id;
  try {
    await host.ensureReady();
    if (req.type === 'parse') {
      doc = null;
      if (localMetricFontFaces.length > 0) {
        unloadLocalFontMetrics(localMetricFontFaces);
        localMetricFontFaces = [];
      }
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's image.
      imageCache.clear();
      // A re-parse starts a fresh document: also drop the shared, per-`getImage`
      // decoded caches (base raster, a:clrChange/duotone recolour, SVG object
      // URLs), symmetric with DocxDocument.destroy(). The worker's `getImage`
      // closure is a stable module-level identity, so without this a new document
      // sharing a zip path (e.g. word/media/image1.png) would be served the
      // previous file's decoded bitmap, and the GPU/URL handles would linger past
      // the LRU cap. Symmetric across docx/pptx/xlsx render workers (issue #781).
      dropBitmapCacheByPath(getImage);
      dropColorReplacedCache(getImage);
      dropSvgImageCache(getImage);
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const bytes = new Uint8Array(req.data);
      // Construction + `parse()` run under `host.run` so a trap in EITHER poisons
      // + recycles the instance (and frees the archive). `setArchive` frees any
      // prior handle first — the re-parse dispose. `parse()` throws on
      // parse/serialize failure (Result<Vec<u8>, JsValue>); the outer try/catch
      // converts a graceful failure into an error response. Render mode consumes
      // the model in-worker, so decode + parse it here (one decode, no
      // passthrough).
      const parsedModel = host.run(() => {
        const archive = new DocxArchive(bytes, max);
        host.setArchive(archive);
        return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel;
      });
      const model = normalizeInternalDocumentModel(parsedModel).document;
      let googleFaces: FontFace[] = [];
      if (req.useGoogleFonts) {
        // Pagination measures text, so fonts must land BEFORE computePages —
        // same ordering the main-mode load() guarantees.
        googleFaces = await preloadGoogleFonts(
          docxFontPreloadNames(model),
          DOCX_GOOGLE_FONTS,
        );
      }
      // ECMA-376 §17.8.1 / §17.8.3 — register embedded fonts into the worker's
      // FontFaceSet (self.fonts) before pagination measures text. Bytes are read
      // straight from the retained archive (extract_image reads any zip entry).
      let embeddedFaces: FontFace[] = [];
      if (model.embeddedFonts?.length) {
        embeddedFaces = await loadEmbeddedFonts(model, async (p) => {
          const loaded = host.archive;
          if (!loaded) throw new Error('No docx loaded');
          return new Uint8Array(host.run(() => loaded.extract_image(p))).slice();
        });
      }
      const localMetrics = await loadDocxLocalFontMetrics(model);
      localMetricFontFaces = localMetrics.faces;
      const layoutServices = createLayoutServices(model, {
        localMetrics: localMetrics.metrics,
        useGoogleFonts: !!req.useGoogleFonts,
        embeddedFaces,
        googleFaces,
      });
      const retainedErrorLayout = model.parseError
        ? deepFreezeDocumentLayout(layoutParseErrorPage(
            model.parseError,
            { widthPt: model.section.pageWidth, heightPt: model.section.pageHeight },
            layoutServices.text,
          )) as DocumentLayout
        : null;
      doc = { model, layoutServices, defaultCurrentDateMs: req.defaultCurrentDateMs, retainedErrorLayout };
      pages = paginateDocument(model, layoutServices, { currentDateMs: req.defaultCurrentDateMs });
      // ECMA-376 §17.6.13 / §17.6.11 / §17.6.20 — per-page size from each page's
      // stamped frame (its page-meta for an empty parity page). A vertical
      // section paginates on the SWAPPED logical geometry, so
      // `physicalPageSizeForPage` — the resolver shared with
      // `DocxDocument.pageSize` — un-swaps the stamped dims by the PAGE's OWN
      // direction (per-section, issue #1000) back to the PHYSICAL page box the
      // meta reports (identity for horizontal pages).
      const paginated = pages;
      const pageSizes = paginated.map((_els, i) => physicalPageSizeForPage(paginated, i, model.section));
      const meta: DocumentMeta = {
        pageCount: pages.length,
        comments: model.comments ?? [],
        footnotes: model.footnotes ?? [],
        endnotes: model.endnotes ?? [],
        pageSizes,
        bookmarkPages: [...buildBookmarkPageMap(pages)],
      };
      post({ type: 'parsedMeta', id, meta });
      return;
    }
    if (req.type === 'renderPage') {
      if (!doc || !pages) throw new Error('Document not loaded');
      const canvas = new OffscreenCanvas(1, 1); // renderer resizes it
      // IX6 — collect the run geometry the same render emits so the main thread
      // can build its selection / find overlay without a second render. The
      // callback runs worker-side; only the resulting plain array crosses back.
      const runs: DocxTextRunInfo[] = [];
      await renderDocumentToCanvas(doc.model, canvas, req.pageIndex, {
        ...req.opts,
        totalPages: pages.length,
        prebuiltPages: pages,
        fetchImage: getImage,
        onTextRun: (r) => runs.push(r),
        layoutServices: doc.layoutServices,
        retainedLayout: doc.retainedErrorLayout ?? undefined,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
      });
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'pageRendered', id, bitmap, runs }, [bitmap]);
      return;
    }
    if (req.type === 'collectRuns') {
      // IX6 — render a page purely to harvest its text-run geometry (find scans
      // every page). The bitmap is discarded worker-side; only `runs` crosses
      // the wire. Same renderer / prebuilt pagination as `renderPage`, so the
      // geometry is identical to what a `renderPage` of the same page would draw.
      if (!doc || !pages) throw new Error('Document not loaded');
      const canvas = new OffscreenCanvas(1, 1);
      const runs: DocxTextRunInfo[] = [];
      await renderDocumentToCanvas(doc.model, canvas, req.pageIndex, {
        ...req.opts,
        totalPages: pages.length,
        prebuiltPages: pages,
        fetchImage: getImage,
        onTextRun: (r) => runs.push(r),
        layoutServices: doc.layoutServices,
        retainedLayout: doc.retainedErrorLayout ?? undefined,
        defaultCurrentDateMs: doc.defaultCurrentDateMs,
      });
      post({ type: 'runsCollected', id, runs });
      return;
    }
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      const archive = host.archive;
      if (!archive) throw new Error('No docx loaded');
      const raw = host.run(() => archive.extract_image(req.path));
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      const archive = host.archive;
      if (!archive) throw new Error('No docx loaded');
      const markdown = host.run(() => archive.to_markdown());
      post({ type: 'markdownRendered', id, markdown });
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
