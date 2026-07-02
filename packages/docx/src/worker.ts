import init, { parse_docx, extract_image } from './wasm/docx_parser.js';
import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';

let initPromise: Promise<unknown> | null = null;
// The buffer is transferred into the worker on `parse` (the main thread's copy
// is neutered), so the worker is its rightful owner. Retain it so a later
// `extractImage` can read media bytes by zip path without re-sending the file.
let currentBuffer: Uint8Array | null = null;
let currentMaxZipEntryBytes: bigint | undefined;

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
      currentMaxZipEntryBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      currentBuffer = new Uint8Array(req.data);
      // `parse_docx` returns the model as UTF-8 JSON bytes on success and throws
      // a JS Error on parse/serialize failure (Result<Vec<u8>, JsValue>),
      // matching pptx/xlsx. The throw is caught by the outer try/catch below, so
      // no error-field probe is needed here. wasm-bindgen hands back a fresh
      // Uint8Array (a copy of the Rust Vec), so its buffer is exclusively ours:
      // forward it to the main thread as a transferable — no clone, no decode
      // here. The single decode + JSON.parse happens once, on the main thread.
      const json = parse_docx(currentBuffer, currentMaxZipEntryBytes);
      const documentJson = json.buffer as ArrayBuffer;
      const res: WorkerResponse = { type: 'parsed', id, documentJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [
        documentJson,
      ]);
      return;
    }
    if (req.type === 'extractImage') {
      if (!currentBuffer) throw new Error('No docx loaded');
      const bytes = extract_image(currentBuffer, req.path, currentMaxZipEntryBytes);
      const copy = new Uint8Array(bytes).slice().buffer;
      const res: WorkerResponse = { type: 'imageExtracted', id, bytes: copy };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(res, [copy]);
      return;
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
