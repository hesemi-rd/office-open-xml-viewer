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
import { decodeDataUrl, preloadGoogleFonts, WasmParserHost } from '@silurus/ooxml-core';
import type { DocxDocumentModel, PaginatedBodyElement } from './types';
import { paginateDocument, renderDocumentToCanvas, physicalPageSizePt } from './renderer';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import { loadEmbeddedFonts } from './embedded-fonts';
import type { RenderWorkerRequest, RenderWorkerResponse, DocumentMeta } from './worker-protocol';

// RB6: self-poison + auto-respawn. A trap during parse (or an in-worker image /
// embedded-font read) recycles the instance so the next document renders on
// clean linear memory. The host owns the `DocxArchive` handle (`host.archive`).
const host = new WasmParserHost<DocxArchive>(init, {
  freeArchive: (a) => a.free(),
  // RB6 recovery must re-instantiate, not re-`init` (a no-op against the
  // wasm-bindgen singleton). `reinit` forces fresh linear memory after a trap.
  reinit,
});
let doc: DocxDocumentModel | null = null;
let pages: PaginatedBodyElement[][] | null = null;
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
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's image.
      imageCache.clear();
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
      doc = host.run(() => {
        const archive = new DocxArchive(bytes, max);
        host.setArchive(archive);
        return JSON.parse(new TextDecoder().decode(archive.parse())) as DocxDocumentModel;
      });
      if (req.useGoogleFonts) {
        // Pagination measures text, so fonts must land BEFORE computePages —
        // same ordering the main-mode load() guarantees.
        await preloadGoogleFonts(
          docxFontPreloadNames(doc),
          DOCX_GOOGLE_FONTS,
        );
      }
      // ECMA-376 §17.8.1 / §17.8.3 — register embedded fonts into the worker's
      // FontFaceSet (self.fonts) before pagination measures text. Bytes are read
      // straight from the retained archive (extract_image reads any zip entry).
      if (doc.embeddedFonts?.length) {
        await loadEmbeddedFonts(doc, async (p) => {
          const loaded = host.archive;
          if (!loaded) throw new Error('No docx loaded');
          return new Uint8Array(host.run(() => loaded.extract_image(p))).slice();
        });
      }
      pages = paginateDocument(doc);
      // ECMA-376 §17.6.13 / §17.6.11 — per-page size from each page's first
      // element's stamped `sectionGeom` (body-level fallback for an empty page).
      const model = doc;
      const pageSizes = pages.map((els) => {
        const g = els[0]?.sectionGeom;
        // A vertical (tbRl) section paginates on the SWAPPED logical geometry, so
        // un-swap the stamped dims back to the PHYSICAL page box the meta reports
        // (identity for horizontal docs).
        return physicalPageSizePt(
          model.section,
          g?.pageWidth ?? model.section.pageWidth,
          g?.pageHeight ?? model.section.pageHeight,
        );
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
