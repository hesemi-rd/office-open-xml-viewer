import type {
  ViewportRange,
  RenderViewportOptions,
  WorkerResponse,
  ParsedWorkbook,
  Worksheet,
} from './types.js';

/** Serializable subset of RenderViewportOptions: drop the callback, the image
 *  cache, and the `fetchImage` loader (all non-cloneable; the worker owns its
 *  own cache and supplies its own in-worker fetchImage). */
export type WireRenderViewportOptions = Omit<
  RenderViewportOptions,
  'onTextRun' | 'loadedImages' | 'fetchImage'
>;

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
  // load-bearing here; `sheetName` is carried for shape-compat with the
  // main-mode message getWorksheet posts but is ignored (the worker derives the
  // sheet name from its own `workbook`). No `data`: like main-mode `parseSheet`,
  // the buffer retained at `parse` is reused — never re-sent per sheet.
  | { type: 'parseSheet'; id: number; sheetIndex: number; sheetName?: string; maxZipEntryBytes?: number }
  | { type: 'renderViewport'; id: number; sheetIndex: number; viewport: ViewportRange; opts: WireRenderViewportOptions }
  // Worker render mode decodes images in-worker via a getImage closure; this arm
  // exists only for protocol parity with worker.ts (so a stray extractImage
  // never hangs). The render worker reads bytes straight from its retained
  // rawData.
  | { type: 'extractImage'; id: number; path: string }
  | { type: 'toMarkdown'; id: number };

export type RenderWorkerResponse =
  // `imageExtracted` / `error` are reused from WorkerResponse. `parsed` and
  // `parsedSheet` are NOT reused: the parse-only worker returns those models as
  // transferred JSON bytes, but the render worker has already decoded them
  // worker-side (it consumes them to render), so it sends the objects back to
  // the proxy as structured clones — no re-serialization. The light,
  // workbook-level ParsedWorkbook keeps synchronous getters (sheetNames,
  // tabColors, …) working; per-sheet data stays worker-side and is parsed on
  // demand.
  | Exclude<WorkerResponse, { type: 'parsed' } | { type: 'parsedSheet' }>
  | { type: 'parsed'; id: number; workbook: ParsedWorkbook }
  | { type: 'parsedSheet'; id: number; worksheet: Worksheet }
  | { type: 'viewportRendered'; id: number; bitmap: ImageBitmap };
