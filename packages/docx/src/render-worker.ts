/**
 * Render-capable worker entry: parse → font preload → paginate, all
 * worker-side; renders pages into an OffscreenCanvas and replies with
 * transferable ImageBitmaps. Used by DocxDocument.load(src, { mode: 'worker' });
 * the slim parse-only worker.ts stays untouched so main-mode users pay no
 * bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 */
import init, { parse_docx, extract_image } from './wasm/docx_parser.js';
import { decodeDataUrl, preloadGoogleFonts } from '@silurus/ooxml-core';
import type { DocxDocumentModel, PaginatedBodyElement } from './types';
import { paginateDocument, renderDocumentToCanvas } from './renderer';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import type { RenderWorkerRequest, RenderWorkerResponse, DocumentMeta } from './worker-protocol';

let initPromise: Promise<unknown> | null = null;
let doc: DocxDocumentModel | null = null;
let pages: PaginatedBodyElement[][] | null = null;
// The buffer is transferred into the worker on `parse` (main thread's copy
// neutered), so the worker owns it. Retained so the in-worker `getImage`
// closure can read image bytes by zip path straight from it (no transfer).
let currentBuffer: Uint8Array | null = null;
let currentMaxZipEntryBytes: bigint | undefined;
const imageCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** In-worker image-byte loader (twin of pptx's render-worker `getImage`). The
 *  renderer's `fetchImage` routes here in worker mode, so image bytes are
 *  decoded straight from the retained buffer with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    if (!currentBuffer) throw new Error('No docx loaded');
    const bytes = extract_image(currentBuffer, path, currentMaxZipEntryBytes);
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
      currentMaxZipEntryBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      currentBuffer = new Uint8Array(req.data);
      const parsed = JSON.parse(parse_docx(currentBuffer, currentMaxZipEntryBytes));
      if (parsed.error) throw new Error(`Parse error: ${parsed.error}`);
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
      const meta: DocumentMeta = {
        pageCount: pages.length,
        comments: doc.comments ?? [],
        footnotes: doc.footnotes ?? [],
        endnotes: doc.endnotes ?? [],
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
      // read straight from the retained buffer (no mime needed for a byte
      // transfer).
      if (!currentBuffer) throw new Error('No docx loaded');
      const raw = extract_image(currentBuffer, req.path, currentMaxZipEntryBytes);
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
