import init, { XlsxArchive } from './wasm/xlsx_parser.js';
import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types.js';

let initPromise: Promise<unknown> | null = null;
// An `XlsxArchive` handle over the opened zip. `new XlsxArchive(bytes, max)`
// copies the file into WASM ONCE and scans the central directory ONCE; the
// workbook / sharedStrings / theme parts are then parsed ONCE and reused on every
// `parseSheet` (the D3 win — a sheet switch previously re-parsed all of them).
// `extractImage` also reads by zip path straight from the retained archive. No
// JS-side buffer is kept alive (the sole copy lives in WASM). Freed + replaced on
// a re-parse.
let archive: XlsxArchive | null = null;

/** Free the current handle (if any) and null it out — double-free / UAF guard. */
function disposeArchive(): void {
  if (archive) {
    archive.free();
    archive = null;
  }
}

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.type === 'init') {
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Every non-init request carries a correlation id that must be echoed back so
  // the client can route the response to the right pending promise.
  const id = req.id;
  try {
    await initPromise;
    if (req.type === 'parse') {
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // Constructing the handle copies `req.data` into WASM; after that the
      // worker holds no reference to those bytes (memory is not doubled). Replace
      // any prior handle first so re-parsing a new file frees the old archive.
      disposeArchive();
      const bytes = new Uint8Array(req.data);
      archive = new XlsxArchive(bytes, max);
      // `parse()` returns the workbook index as UTF-8 JSON bytes (Result<Vec<u8>,
      // JsValue>). wasm-bindgen hands back a fresh Uint8Array that owns its
      // buffer, so forward it to main as a transferable — no clone, no decode
      // here. The single decode + JSON.parse happens on main.
      const json = archive.parse();
      const workbookJson = json.buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'parsed', id, workbookJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [
        workbookJson,
      ]);
    } else if (req.type === 'parseSheet') {
      // Parse straight off the retained archive, reusing its cached workbook /
      // sharedStrings / theme parts — the whole file no longer crosses the worker
      // boundary and those shared parts are no longer re-parsed on every sheet
      // switch. A `parseSheet` before any `parse` has no archive to work with:
      // that is a protocol violation, so fail loudly.
      if (!archive) {
        throw new Error('parseSheet before parse: no archive retained');
      }
      // `parse_sheet` also returns UTF-8 JSON bytes; forward its transferable
      // buffer to main the same way (single decode + parse on main).
      const json = archive.parse_sheet(req.sheetIndex, req.sheetName);
      const worksheetJson = json.buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'parsedSheet', id, worksheetJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [
        worksheetJson,
      ]);
    } else if (req.type === 'extractImage') {
      if (!archive) throw new Error('No xlsx loaded');
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `.buffer` is a full-span, non-WASM-backed ArrayBuffer we own outright —
      // transfer it directly. A second `new Uint8Array(bytes).slice()` would just
      // re-copy the whole entry for nothing.
      const out = archive.extract_image(req.path).buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'imageExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [out]);
    } else if (req.type === 'toMarkdown') {
      if (!archive) throw new Error('No xlsx loaded');
      // Project the already-opened handle to markdown (no re-copy of the file,
      // no re-scan of the central directory). A plain string has no transferable
      // backing, so it is posted by structured clone like any other value.
      const markdown = archive.to_markdown();
      const res: WorkerResponse = { type: 'markdownRendered', id, markdown };
      self.postMessage(res);
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
