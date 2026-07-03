import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';
import init, { PptxArchive } from './wasm/pptx_parser.js';

let initPromise: Promise<unknown> | null = null;
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

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    // Retain the init promise (docx/xlsx pattern) rather than a `ready` flag +
    // handshake. Every request below `await`s it, so a REJECTED init rejects the
    // request (the catch posts an `error` response the bridge turns into a
    // rejected `load()`), never a silent hang on a main-side `ready` wait.
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Echo the correlation id so the client routes the response to the right
  // pending promise (id correlation, not response-type matching).
  const id = req.id;
  try {
    await initPromise;
    if (req.kind === 'parse') {
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
      const msg: WorkerResponse = { kind: 'parsed', id, presentationJson };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [
        presentationJson,
      ]);
      return;
    }

    if (req.kind === 'extractMedia') {
      if (!archive) throw new Error('No pptx loaded');
      // wasm-bindgen already hands back a fresh, standalone Uint8Array here (its
      // glue does `getArrayU8FromWasm0(ptr,len).slice()` then frees the Rust Vec),
      // so `bytes.buffer` is a full-span, non-WASM-backed ArrayBuffer we own
      // outright — transfer it directly. A second `new Uint8Array(bytes).slice()`
      // would just re-copy the whole entry for nothing.
      const out = archive.extract_media(req.path).buffer as ArrayBuffer;
      const msg: WorkerResponse = { kind: 'mediaExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [out]);
      return;
    }

    if (req.kind === 'extractImage') {
      if (!archive) throw new Error('No pptx loaded');
      // See extractMedia above: the extracted Uint8Array already owns a
      // standalone full-span buffer, so transfer it without a second copy.
      const out = archive.extract_image(req.path).buffer as ArrayBuffer;
      const msg: WorkerResponse = { kind: 'imageExtracted', id, bytes: out };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [out]);
      return;
    }

    if (req.kind === 'toMarkdown') {
      if (!archive) throw new Error('No pptx loaded');
      // Project the already-opened handle to markdown (no re-copy of the file,
      // no re-scan of the central directory). A plain string has no transferable
      // backing, so it is posted by structured clone like any other value.
      const markdown = archive.to_markdown();
      const msg: WorkerResponse = { kind: 'markdownRendered', id, markdown };
      self.postMessage(msg);
      return;
    }
  } catch (err) {
    const msg: WorkerResponse = {
      kind: 'error',
      id,
      message: err instanceof Error ? err.message : String(err),
    };
    self.postMessage(msg);
  }
};
