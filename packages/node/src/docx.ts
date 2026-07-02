import type { DocxDocumentModel } from '@silurus/ooxml-docx';
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

let initialized = false;

function ensureInit(): void {
  if (initialized) return;
  const wasmPath = resolveWasm(import.meta.url, '../../docx/src/wasm/docx_parser_bg.wasm');
  loadWasmModule(docxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown }, wasmPath);
  initialized = true;
}

/** Parse a `.docx` archive in Node and return the same `DocxDocumentModel` the
 *  browser path produces. */
export function parseDocx(buffer: ArrayBuffer | Uint8Array | Buffer): DocxDocumentModel {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  // `parse_docx` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once. Matches the browser main-thread receiver.
  const json = (docxWasm as unknown as { parse_docx: (b: Uint8Array) => Uint8Array }).parse_docx(
    bytes,
  );
  return JSON.parse(new TextDecoder().decode(json)) as DocxDocumentModel;
}
