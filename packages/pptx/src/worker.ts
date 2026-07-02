import type { WorkerRequest, WorkerResponse } from './types';
import init, { parse_pptx, extract_media, extract_image } from './wasm/pptx_parser.js';

let ready = false;
let currentBuffer: Uint8Array | null = null;
let currentMaxZipEntryBytes: bigint | undefined;

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
      const bytes = new Uint8Array(req.buffer);
      currentBuffer = bytes;
      currentMaxZipEntryBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // `parse_pptx` returns the model as UTF-8 JSON bytes (Result<Vec<u8>,
      // JsValue>). wasm-bindgen hands back a fresh Uint8Array that owns its
      // buffer, so forward it to the main thread as a transferable — no clone,
      // no decode here. The single decode + JSON.parse happens once, on main.
      const json = parse_pptx(bytes, currentMaxZipEntryBytes);
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
    if (!currentBuffer) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'No pptx loaded' };
      self.postMessage(msg);
      return;
    }
    try {
      const bytes = extract_media(currentBuffer, req.path, currentMaxZipEntryBytes);
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
    if (!currentBuffer) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'No pptx loaded' };
      self.postMessage(msg);
      return;
    }
    try {
      const bytes = extract_image(currentBuffer, req.path, currentMaxZipEntryBytes);
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
