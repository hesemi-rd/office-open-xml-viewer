import type { Presentation } from '@silurus/ooxml-pptx';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

let initialized = false;

function ensureInit(): void {
  if (initialized) return;
  // The wasm-pack `--target web` JS module exports `initSync` and the
  // parser functions. Locate the sibling `.wasm` file in the pptx package
  // and feed its bytes into `initSync` so the module is fully linked
  // before the first `parse_pptx` call.
  const wasmPath = resolveWasm(import.meta.url, '../../pptx/src/wasm/pptx_parser_bg.wasm');
  loadWasmModule(pptxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown }, wasmPath);
  initialized = true;
}

/** Parse a `.pptx` archive in Node and return the same `Presentation` model
 *  the browser path produces. Synchronous WASM init is performed on first call
 *  and cached for subsequent invocations. */
export function parsePptx(buffer: ArrayBuffer | Uint8Array | Buffer): Presentation {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  const json = (pptxWasm as unknown as { parse_pptx: (b: Uint8Array) => string }).parse_pptx(bytes);
  return JSON.parse(json) as Presentation;
}

/** Extract raw bytes for a single media entry (e.g. `ppt/media/image1.png`)
 *  from the source archive. Mirrors `extract_media` on the browser worker. */
export function extractMedia(buffer: ArrayBuffer | Uint8Array | Buffer, path: string): Uint8Array {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  return (pptxWasm as unknown as { extract_media: (b: Uint8Array, p: string) => Uint8Array }).extract_media(bytes, path);
}

/** Extract raw bytes for a single embedded image entry (e.g.
 *  `ppt/media/image1.png`) from the source archive. Mirrors `extract_image`
 *  on the browser worker (twin of {@link extractMedia}); pictures and blip
 *  fills now carry only zip paths, so the render path reads image bytes lazily
 *  through this. `maxZipEntryBytes` mirrors the worker's per-entry guard and is
 *  optional (no cap when omitted). */
export function extractImage(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  path: string,
  maxZipEntryBytes?: number,
): Uint8Array {
  ensureInit();
  const bytes =
    buffer instanceof Uint8Array
      ? buffer
      : new Uint8Array(buffer as ArrayBuffer);
  return (
    pptxWasm as unknown as {
      extract_image: (b: Uint8Array, p: string, max?: bigint) => Uint8Array;
    }
  ).extract_image(
    bytes,
    path,
    typeof maxZipEntryBytes === 'number' && maxZipEntryBytes > 0
      ? BigInt(maxZipEntryBytes)
      : undefined,
  );
}
