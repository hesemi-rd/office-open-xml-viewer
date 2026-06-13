/**
 * Render-capable worker entry: parse → font preload → paginate, all
 * worker-side; renders pages into an OffscreenCanvas and replies with
 * transferable ImageBitmaps. Used by DocxDocument.load(src, { mode: 'worker' });
 * the slim parse-only worker.ts stays untouched so main-mode users pay no
 * bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 */
import init, { parse_docx } from './wasm/docx_parser.js';
import { decodeDataUrl, preloadGoogleFonts } from '@silurus/ooxml-core';
import type { DocxDocumentModel, PaginatedBodyElement } from './types';
import { paginateDocument, renderDocumentToCanvas } from './renderer';
import { DOCX_GOOGLE_FONTS, docxFontPreloadNames } from './google-fonts';
import type { RenderWorkerRequest, RenderWorkerResponse, DocumentMeta } from './worker-protocol';

let initPromise: Promise<unknown> | null = null;
let doc: DocxDocumentModel | null = null;
let pages: PaginatedBodyElement[][] | null = null;

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

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
      const maxBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const parsed = JSON.parse(parse_docx(new Uint8Array(req.data), maxBytes));
      if (parsed.error) throw new Error(`Parse error: ${parsed.error}`);
      doc = parsed as DocxDocumentModel;
      if (req.useGoogleFonts) {
        // Pagination measures text, so fonts must land BEFORE computePages —
        // same ordering the main-mode load() guarantees.
        await preloadGoogleFonts(
          docxFontPreloadNames(doc.majorFont, doc.minorFont),
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
      });
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'pageRendered', id, bitmap }, [bitmap]);
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
