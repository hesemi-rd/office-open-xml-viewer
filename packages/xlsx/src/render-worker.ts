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
import init, { parse_xlsx, parse_sheet } from './wasm/xlsx_parser.js';
import { decodeDataUrl, preloadGoogleFonts } from '@silurus/ooxml-core';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { XLSX_GOOGLE_FONTS, xlsxFontPreloadNames } from './google-fonts.js';
import type { ParsedWorkbook, Worksheet } from './types.js';
import type { RenderWorkerRequest, RenderWorkerResponse } from './worker-protocol.js';

let initPromise: Promise<unknown> | null = null;
let workbook: ParsedWorkbook | null = null;
let rawData: ArrayBuffer | null = null;
let maxZipEntryBytes: bigint | undefined;
let fontsLoaded: Promise<void> = Promise.resolve();
const sheetCache = new Map<number, Worksheet>();
const imageCache = new Map<string, CanvasImageSource>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** Lazily parse one worksheet directly via WASM and cache it — the worker-side
 *  equivalent of XlsxWorkbook.getWorksheet (same `parse_sheet` call, same
 *  cache). Shared by the `renderViewport` handler. */
function parseSheetLocally(sheetIndex: number): Worksheet {
  const cached = sheetCache.get(sheetIndex);
  if (cached) return cached;
  if (!workbook || !rawData) throw new Error('Workbook not loaded');
  const sheetMeta = workbook.workbook.sheets[sheetIndex];
  if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);
  const json = parse_sheet(new Uint8Array(rawData), sheetIndex, sheetMeta.name, maxZipEntryBytes);
  const ws = JSON.parse(json) as Worksheet;
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
      maxZipEntryBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const json = parse_xlsx(new Uint8Array(req.data), maxZipEntryBytes);
      workbook = JSON.parse(json) as ParsedWorkbook;
      rawData = req.data;
      if (req.useGoogleFonts) {
        // Mirror XlsxWorkbook._load exactly: queue Google Fonts substitutes for
        // every styled font name, plus the generic Arabic fallbacks. Fonts must
        // land before rendering (which measures text), so we keep the promise
        // and await it in the renderViewport handler.
        fontsLoaded = preloadGoogleFonts(xlsxFontPreloadNames(workbook.styles), XLSX_GOOGLE_FONTS);
      }
      post({ type: 'parsed', id, workbook });
      return;
    }
    if (req.type === 'parseSheet') {
      // Worker-side equivalent of the slim worker.ts `parseSheet` arm, so
      // XlsxWorkbook.getWorksheet / resolveValidationList resolve in worker
      // mode instead of hanging forever. Reuses parseSheetLocally (which parses
      // from the worker's stored rawData and cache); the main-mode message's
      // `data` / `sheetName` fields are ignored. Reply shape matches worker.ts.
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
        req.opts,
      );
      const bitmap = canvas.transferToImageBitmap();
      post({ type: 'viewportRendered', id, bitmap }, [bitmap]);
      return;
    }
  } catch (err) {
    post({ type: 'error', id, message: err instanceof Error ? err.message : String(err) });
  }
};
