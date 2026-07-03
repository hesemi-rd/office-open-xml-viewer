/**
 * Render-capable worker entry: parse → font preload → paginate, all
 * worker-side; renders pages into an OffscreenCanvas and replies with
 * transferable ImageBitmaps. Used by DocxDocument.load(src, { mode: 'worker' });
 * the slim parse-only worker.ts stays untouched so main-mode users pay no
 * bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 */
import init, { DocxArchive } from './wasm/docx_parser.js';
import { decodeDataUrl, preloadGoogleFonts } from '@silurus/ooxml-core';
import type { DocxDocumentModel, PaginatedBodyElement } from './types';
import { paginateDocument, renderDocumentToCanvas } from './renderer';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import type { RenderWorkerRequest, RenderWorkerResponse, DocumentMeta } from './worker-protocol';

let initPromise: Promise<unknown> | null = null;
let doc: DocxDocumentModel | null = null;
let pages: PaginatedBodyElement[][] | null = null;
// A `DocxArchive` handle over the opened zip. `new DocxArchive(bytes, max)`
// copies the file into WASM ONCE and opens it ONCE; the in-worker `getImage`
// closure then reads image bytes by zip path straight from the retained archive
// (no transfer, no re-open, no JS-side buffer kept alive). Freed + replaced on a
// re-parse.
let archive: DocxArchive | null = null;
const imageCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** Free the current handle (if any) and null it out — double-free / UAF guard. */
function disposeArchive(): void {
  if (archive) {
    archive.free();
    archive = null;
  }
}

/** In-worker image-byte loader (twin of pptx's render-worker `getImage`). The
 *  renderer's `fetchImage` routes here in worker mode, so image bytes are
 *  decoded straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    if (!archive) throw new Error('No docx loaded');
    const bytes = archive.extract_image(path);
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageCache.set(path, p);
  return p;
}

self.onmessage = async (e: MessageEvent<RenderWorkerRequest>) => {
  const req = e.data;
  if (req.type === 'init') {
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }
  const id = req.id;
  try {
    await initPromise;
    if (req.type === 'parse') {
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's image.
      imageCache.clear();
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // Constructing the handle copies `req.data` into WASM; the worker then
      // holds no reference to the transferred bytes (memory is not doubled).
      // Replace any prior handle first so a re-parse frees the old archive.
      disposeArchive();
      const bytes = new Uint8Array(req.data);
      archive = new DocxArchive(bytes, max);
      // `parse()` throws on parse/serialize failure (Result<Vec<u8>, JsValue>);
      // the outer try/catch converts it into an error response, so the bytes here
      // are always the success model — no error-field probe. Render mode consumes
      // the model in-worker, so decode + parse it here (one decode, no
      // passthrough).
      const parsed = JSON.parse(new TextDecoder().decode(archive.parse()));
      doc = parsed as DocxDocumentModel;
      if (req.useGoogleFonts) {
        // Pagination measures text, so fonts must land BEFORE computePages —
        // same ordering the main-mode load() guarantees.
        await preloadGoogleFonts(
          docxFontPreloadNames(doc),
          DOCX_GOOGLE_FONTS,
        );
      }
      pages = paginateDocument(doc);
      // ECMA-376 §17.6.13 / §17.6.11 — per-page size from each page's first
      // element's stamped `sectionGeom` (body-level fallback for an empty page).
      const model = doc;
      const pageSizes = pages.map((els) => {
        const g = els[0]?.sectionGeom;
        return {
          widthPt: g?.pageWidth ?? model.section.pageWidth,
          heightPt: g?.pageHeight ?? model.section.pageHeight,
        };
      });
      const meta: DocumentMeta = {
        pageCount: pages.length,
        comments: doc.comments ?? [],
        footnotes: doc.footnotes ?? [],
        endnotes: doc.endnotes ?? [],
        pageSizes,
      };
      post({ type: 'parsedMeta', id, meta });
      return;
    }
    if (req.type === 'renderPage') {
      if (!doc || !pages) throw new Error('Document not loaded');
      const canvas = new OffscreenCanvas(1, 1); // renderer resizes it
      await renderDocumentToCanvas(doc, canvas, req.pageIndex, {
        ...req.opts,
        totalPages: pages.length,
        prebuiltPages: pages,
        fetchImage: getImage,
      });
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'pageRendered', id, bitmap }, [bitmap]);
      return;
    }
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      if (!archive) throw new Error('No docx loaded');
      const raw = archive.extract_image(req.path);
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
    if (req.type === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      if (!archive) throw new Error('No docx loaded');
      const markdown = archive.to_markdown();
      post({ type: 'markdownRendered', id, markdown });
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
