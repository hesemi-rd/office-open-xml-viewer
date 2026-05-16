import type { Document } from '@silurus/ooxml-docx';
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

/** Parse a `.docx` archive in Node and return the same `Document` model the
 *  browser path produces. */
export function parseDocx(buffer: ArrayBuffer | Uint8Array | Buffer): Document {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  const json = (docxWasm as unknown as { parse_docx: (b: Uint8Array) => string }).parse_docx(bytes);
  return JSON.parse(json) as Document;
}
