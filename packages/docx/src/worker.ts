import init, { DocxArchive } from './wasm/docx_parser.js';
import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';

let initPromise: Promise<unknown> | null = null;
// A `DocxArchive` handle over the opened zip: `new DocxArchive(bytes, max)`
// copies the file into WASM ONCE and scans the central directory ONCE, then a
// later `extractImage` reads media by zip path straight from the retained
// archive (no re-copy, no re-open, and no JS-side buffer kept alive — the copy
// lives inside WASM). Held across the worker's lifetime; freed + replaced on a
// re-parse so we never leak the previous document's WASM allocation.
let archive: DocxArchive | null = null;

/** Free the current handle (if any) and null it out. Guards against a double
 *  free / use-after-free: after this the next `extractImage` throws the same
 *  "No docx loaded" error as before a first parse. */
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

  // Echo the correlation id so the client routes the response to the right
  // pending promise (id correlation, not response-type matching).
  const id = req.id;
  try {
    await initPromise;
    if (req.type === 'parse') {
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // Constructing the handle copies `req.data` into WASM; after that the
      // worker holds no reference to the transferred bytes (memory is not
      // doubled — the sole copy lives in WASM linear memory). Replace any prior
      // handle first so re-parsing a new file frees the old archive.
      disposeArchive();
      const bytes = new Uint8Array(req.data);
      archive = new DocxArchive(bytes, max);
      // `parse()` returns the model as UTF-8 JSON bytes on success and throws a
      // JS Error on parse/serialize failure (Result<Vec<u8>, JsValue>). The throw
      // is caught by the outer try/catch below. wasm-bindgen hands back a fresh
      // Uint8Array (a copy of the Rust Vec), so its buffer is exclusively ours:
      // forward it to the main thread as a transferable — no clone, no decode
      // here. The single decode + JSON.parse happens once, on the main thread.
      const json = archive.parse();
      const documentJson = json.buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'parsed', id, documentJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [
        documentJson,
      ]);
      return;
    }
    if (req.type === 'extractImage') {
      if (!archive) throw new Error('No docx loaded');
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `.buffer` is a full-span, non-WASM-backed ArrayBuffer we own outright —
      // transfer it directly. A second `new Uint8Array(bytes).slice()` would just
      // re-copy the whole entry for nothing.
      const out = archive.extract_image(req.path).buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'imageExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [out]);
      return;
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
