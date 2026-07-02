import init, { parse_xlsx, parse_sheet, extract_image } from './wasm/xlsx_parser.js';
import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types.js';

let initPromise: Promise<unknown> | null = null;
// The buffer is *copied* into the worker on `parse` (xlsx clones rather than
// transfers, so the main thread keeps its own copy too). Retain it so a later
// `extractImage` can read media bytes by zip path without re-sending the file —
// mirroring how docx/pptx worker.ts keep `currentBuffer`.
let currentBuffer: Uint8Array | null = null;
let currentMaxZipEntryBytes: bigint | undefined;

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
      currentMaxZipEntryBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      currentBuffer = new Uint8Array(req.data);
      const json = parse_xlsx(currentBuffer, currentMaxZipEntryBytes);
      const workbook = JSON.parse(json);
      const res: WorkerResponse = { type: 'parsed', id, workbook };
      self.postMessage(res);
    } else if (req.type === 'parseSheet') {
      // Reuse the buffer retained at `parse` instead of re-receiving it — the
      // whole file no longer crosses the worker boundary on every sheet switch
      // (twin of the `extractImage` reuse below). A `parseSheet` before any
      // `parse` has no buffer to work with: that is a protocol violation, so
      // fail loudly rather than silently returning an empty sheet.
      if (!currentBuffer) {
        throw new Error('parseSheet before parse: no buffer retained');
      }
      const maxBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const json = parse_sheet(currentBuffer, req.sheetIndex, req.sheetName, maxBytes);
      const worksheet = JSON.parse(json);
      const res: WorkerResponse = { type: 'parsedSheet', id, worksheet };
      self.postMessage(res);
    } else if (req.type === 'extractImage') {
      if (!currentBuffer) throw new Error('No xlsx loaded');
      const bytes = extract_image(currentBuffer, req.path, currentMaxZipEntryBytes);
      const copy = new Uint8Array(bytes).slice().buffer;
      const res: WorkerResponse = { type: 'imageExtracted', id, bytes: copy };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [copy]);
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
