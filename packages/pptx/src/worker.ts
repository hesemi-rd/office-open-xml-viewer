import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';
import init, { PptxArchive } from './wasm/pptx_parser.js';

let ready = false;
// A `PptxArchive` handle over the opened zip: `new PptxArchive(bytes, max)`
// copies the file into WASM ONCE and scans the central directory ONCE, then a
// later `extractMedia` / `extractImage` reads by zip path straight from the
// retained archive (no re-copy, no re-open, and no JS-side buffer kept alive —
// the sole copy lives in WASM). Freed + replaced on a re-parse.
let archive: PptxArchive | null = null;

/** Free the current handle (if any) and null it out — double-free / UAF guard. */
function disposeArchive(): void {
  if (archive) {
    archive.free();
    archive = null;
  }
}

async function initWasm(wasmUrl: string) {
  await init(decodeDataUrl(wasmUrl) ?? wasmUrl);
  ready = true;
  const msg: WorkerResponse = { kind: 'ready' };
  self.postMessage(msg);
}

self.onmessage = (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    initWasm(req.wasmUrl).catch((err) => {
      console.error('[pptx-worker] WASM init failed:', err);
    });
    return;
  }

  if (req.kind === 'parse') {
    if (!ready) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'WASM not initialized' };
      self.postMessage(msg);
      return;
    }
    try {
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // Constructing the handle copies `req.buffer` into WASM; after that the
      // worker holds no reference to those bytes (memory is not doubled — the
      // sole copy lives in WASM linear memory). Replace any prior handle first so
      // re-parsing a new file frees the old archive.
      disposeArchive();
      const bytes = new Uint8Array(req.buffer);
      archive = new PptxArchive(bytes, max);
      // `parse()` returns the model as UTF-8 JSON bytes (Result<Vec<u8>,
      // JsValue>). wasm-bindgen hands back a fresh Uint8Array that owns its
      // buffer, so forward it to the main thread as a transferable — no clone,
      // no decode here. The single decode + JSON.parse happens once, on main.
      const json = archive.parse();
      const presentationJson = json.buffer as ArrayBuffer;
      const msg: WorkerResponse = { kind: 'parsed', id: req.id, presentationJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [
        presentationJson,
      ]);
    } catch (err) {
      const msg: WorkerResponse = {
        kind: 'error',
        id: req.id,
        message: err instanceof Error ? err.message : String(err),
      };
      self.postMessage(msg);
    }
    return;
  }

  if (req.kind === 'extractMedia') {
    if (!archive) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'No pptx loaded' };
      self.postMessage(msg);
      return;
    }
    try {
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `bytes.buffer` is a full-span, non-WASM-backed ArrayBuffer we own
      // outright — transfer it directly. A second `new Uint8Array(bytes).slice()`
      // would just re-copy the whole entry for nothing.
      const out = archive.extract_media(req.path).buffer as ArrayBuffer;
      const msg: WorkerResponse = { kind: 'mediaExtracted', id: req.id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [out]);
    } catch (err) {
      const msg: WorkerResponse = {
        kind: 'error',
        id: req.id,
        message: err instanceof Error ? err.message : String(err),
      };
      self.postMessage(msg);
    }
    return;
  }

  if (req.kind === 'extractImage') {
    if (!archive) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'No pptx loaded' };
      self.postMessage(msg);
      return;
    }
    try {
      // See extractMedia above: the extracted Uint8Array already owns a
      // standalone full-span buffer, so transfer it without a second copy.
      const out = archive.extract_image(req.path).buffer as ArrayBuffer;
      const msg: WorkerResponse = { kind: 'imageExtracted', id: req.id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [out]);
    } catch (err) {
      const msg: WorkerResponse = {
        kind: 'error',
        id: req.id,
        message: err instanceof Error ? err.message : String(err),
      };
      self.postMessage(msg);
    }
    return;
  }
};
