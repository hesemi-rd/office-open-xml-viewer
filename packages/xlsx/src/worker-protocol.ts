import type { ViewportRange, RenderViewportOptions, WorkerResponse } from './types.js';

/** Serializable subset of RenderViewportOptions: drop the callback and the
 *  image cache (the worker owns its own). */
export type WireRenderViewportOptions = Omit<RenderViewportOptions, 'onTextRun' | 'loadedImages'>;

// The base `parse` arm from types.ts is intentionally NOT reused: the render
// worker's `parse` carries an extra `useGoogleFonts` flag, and two `parse`
// arms in one union would defeat `type`-based narrowing at use sites. The
// `init` arm is copied verbatim from `WorkerRequest`.
export type RenderWorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number; useGoogleFonts?: boolean }
  // `parseSheet` lets worker-mode XlsxWorkbook.getWorksheet (and the
  // resolveValidationList range path that awaits it) work, mirroring how the
  // pptx render worker handles `extractMedia` for getMedia. The render worker
  // already holds `rawData` + `workbook` from `parse`, so only `sheetIndex` is
  // load-bearing here; `data` / `sheetName` are carried for shape-compat with
  // the main-mode message getWorksheet posts but are ignored (the worker parses
  // from its own `rawData` and derives the sheet name from `workbook`).
  | { type: 'parseSheet'; id: number; data?: ArrayBuffer; sheetIndex: number; sheetName?: string; maxZipEntryBytes?: number }
  | { type: 'renderViewport'; id: number; sheetIndex: number; viewport: ViewportRange; opts: WireRenderViewportOptions };

export type RenderWorkerResponse =
  // `parsed` / `parsedSheet` / `error` are reused from WorkerResponse: xlsx DOES
  // transfer the (light, workbook-level) ParsedWorkbook back to the proxy, so
  // synchronous getters (sheetNames, tabColors, …) keep working; per-sheet data
  // stays worker-side and is parsed there on demand.
  | WorkerResponse
  | { type: 'viewportRendered'; id: number; bitmap: ImageBitmap };
