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

function decodeDataUrl(url: string): ArrayBuffer | null {
  if (!url.startsWith('data:')) return null;
  const comma = url.indexOf(',');
  if (comma === -1) return null;
  const binary = atob(url.slice(comma + 1));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
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
      const bytes = archive.extract_media(req.path);
      const copy = new Uint8Array(bytes).slice().buffer;
      const msg: WorkerResponse = { kind: 'mediaExtracted', id: req.id, bytes: copy };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [copy]);
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
      const bytes = archive.extract_image(req.path);
      const copy = new Uint8Array(bytes).slice().buffer;
      const msg: WorkerResponse = { kind: 'imageExtracted', id: req.id, bytes: copy };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [copy]);
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
