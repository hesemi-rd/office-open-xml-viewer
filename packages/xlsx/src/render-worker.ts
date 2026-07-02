/**
 * Render-capable worker entry: parse → font preload → (lazy) per-sheet parse →
 * render, all worker-side; renders a sheet viewport into an OffscreenCanvas and
 * replies with a transferable ImageBitmap. Used by
 * XlsxWorkbook.load(src, { mode: 'worker' }); the slim parse-only worker.ts
 * stays untouched so main-mode users pay no bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders. A
 * re-`parse` resets all per-document caches so a reused worker never serves
 * stale sheets / images.
 */
import init, { XlsxArchive } from './wasm/xlsx_parser.js';
import { decodeDataUrl, preloadGoogleFonts } from '@silurus/ooxml-core';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames } from './google-fonts.js';
import type { ParsedWorkbook, Worksheet } from './types.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';

let initPromise: Promise<unknown> | null = null;
let workbook: ParsedWorkbook | null = null;
// An `XlsxArchive` handle over the opened zip. `new XlsxArchive(bytes, max)`
// copies the file into WASM ONCE and opens it ONCE; the workbook / sharedStrings
// / theme parts are parsed ONCE and reused on every `parse_sheet` (the D3 win).
// `getImage` also reads bytes by zip path straight from the retained archive. No
// JS-side buffer kept alive. Freed + replaced on a re-parse.
let archive: XlsxArchive | null = null;
let fontsLoaded: Promise<void> = Promise.resolve();
const sheetCache = new Map<number, Worksheet>();
const imageCache = new Map<string, CanvasImageSource | null>();
// Fetched image *bytes* (as Blobs) keyed by zip path. Twin of the docx render
// worker's `imageCache`; kept separate from the decoded-source `imageCache`
// above. Cleared on re-parse so a reused worker never serves a stale file's
// image.
const imageBlobCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** Free the current handle (if any) and null it out — double-free / UAF guard. */
function disposeArchive(): void {
  if (archive) {
    archive.free();
    archive = null;
  }
}

/** In-worker image-byte loader (twin of the docx render-worker `getImage`). The
 *  orchestrator's `fetchImage` routes here in worker mode, so image bytes are
 *  read straight from the retained archive with no main-thread round-trip.
 *  Mime travels on the element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageBlobCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    if (!archive) throw new Error('Workbook not loaded');
    const bytes = archive.extract_image(path);
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageBlobCache.set(path, p);
  return p;
}

/** Lazily parse one worksheet directly via WASM and cache it — the worker-side
 *  equivalent of XlsxWorkbook.getWorksheet (same `parse_sheet` call, same
 *  decode cache). The handle's own shared-part cache means repeat sheet switches
 *  no longer re-parse the workbook / sharedStrings / theme. Shared by the
 *  `renderViewport` handler. */
function parseSheetLocally(sheetIndex: number): Worksheet {
  const cached = sheetCache.get(sheetIndex);
  if (cached) return cached;
  if (!workbook || !archive) throw new Error('Workbook not loaded');
  const sheetMeta = workbook.workbook.sheets[sheetIndex];
  if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);
  // `parse_sheet` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); the
  // render worker consumes the model in-worker, so decode + parse here.
  const json = archive.parse_sheet(sheetIndex, sheetMeta.name);
  const ws = JSON.parse(new TextDecoder().decode(json)) as Worksheet;
  sheetCache.set(sheetIndex, ws);
  return ws;
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
      // A re-parse starts a fresh document: drop any cached sheets / images so
      // we never serve stale data from a previous load.
      sheetCache.clear();
      imageCache.clear();
      imageBlobCache.clear();
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // Constructing the handle copies `req.data` into WASM; the worker then
      // holds no reference to those bytes (memory is not doubled). Replace any
      // prior handle first so a re-parse frees the old archive.
      disposeArchive();
      archive = new XlsxArchive(new Uint8Array(req.data), max);
      // `parse()` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
      // parse the workbook index here (consumed in-worker, then a light copy is
      // sent to the proxy as an object).
      workbook = JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook;
      if (req.useGoogleFonts) {
        // Mirror XlsxWorkbook._load exactly: queue Google Fonts substitutes for
        // every styled font name, plus the generic Arabic fallbacks. Fonts must
        // land before rendering (which measures text), so we keep the promise
        // and await it in the renderViewport handler.
        fontsLoaded = preloadGoogleFonts(xlsxFontPreloadNames(workbook), XLSX_GOOGLE_FONTS);
      }
      post({ type: 'parsed', id, workbook });
      return;
    }
    if (req.type === 'parseSheet') {
      // Worker-side equivalent of the slim worker.ts `parseSheet` arm, so
      // XlsxWorkbook.getWorksheet / resolveValidationList resolve in worker
      // mode instead of hanging forever. Reuses parseSheetLocally (which parses
      // from the worker's stored rawData and cache); the main-mode message's
      // `sheetName` field is ignored. Reply shape matches worker.ts.
      if (!workbook) throw new Error('Workbook not loaded');
      const worksheet = parseSheetLocally(req.sheetIndex);
      post({ type: 'parsedSheet', id, worksheet });
      return;
    }
    if (req.type === 'renderViewport') {
      if (!workbook) throw new Error('Workbook not loaded');
      await fontsLoaded;
      const ws = parseSheetLocally(req.sheetIndex);
      const canvas = new OffscreenCanvas(1, 1); // orchestrator resizes it
      await renderWorksheetViewport(
        { ws, styles: workbook.styles, imageCache },
        canvas,
        req.viewport,
        // Supply the in-worker byte loader so embedded images decode straight
        // from the retained rawData (no main-thread round-trip).
        { ...req.opts, fetchImage: getImage },
      );
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'viewportRendered', id, bitmap }, [bitmap]);
      return;
    }
    if (req.type === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this arm exists only for protocol parity with worker.ts. Raw bytes are
      // read straight from the retained archive (no mime needed for a byte
      // transfer).
      if (!archive) throw new Error('Workbook not loaded');
      const raw = archive.extract_image(req.path);
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ type: 'imageExtracted', id, bytes }, [bytes]);
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
