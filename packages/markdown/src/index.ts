/**
 * Convert .pptx / .docx / .xlsx to GitHub-flavoured markdown via the
 * workspace WASM parsers. The underlying `pptx_to_markdown` /
 * `docx_to_markdown` / `xlsx_to_markdown` functions are wasm_bindgen exports
 * that call the Rust markdown projection directly (no JSON round-trip).
 *
 * Usage modes:
 *
 *   - **Node / CLI**: call `initFromFile()` once at startup, then pass file
 *     bytes (read with `readFileSync`) to the matching `*ToMarkdown`
 *     function. The CLI binary at `bin/ooxml-md.mjs` does this.
 *
 *   - **Browser bundlers (Vite / webpack)**: import the `.wasm` URL via the
 *     bundler-specific syntax (`?url` in Vite), fetch it, and call
 *     `initFromBytes()`. Then call `*ToMarkdown` on a `File.arrayBuffer()`.
 */

// @ts-ignore — wasm-pack JS shim, exports typed individually below
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
// @ts-ignore
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
// @ts-ignore
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';

type WasmModule = {
  initSync: (init: { module: WebAssembly.Module }) => unknown;
};

const pptxState = { initialized: false };
const docxState = { initialized: false };
const xlsxState = { initialized: false };

function syncInit(mod: WasmModule, bytes: Uint8Array, state: { initialized: boolean }): void {
  if (state.initialized) return;
  // The DOM TS lib types `WebAssembly.Module(...)` as `BufferSource =
  // ArrayBufferView<ArrayBuffer> | ArrayBuffer`, which excludes
  // `Uint8Array<SharedArrayBuffer>`. We always pass non-shared bytes here
  // (read from `readFileSync` or `fetch().arrayBuffer()`), so the cast is
  // safe.
  const module = new WebAssembly.Module(bytes as unknown as ArrayBuffer);
  mod.initSync({ module });
  state.initialized = true;
}

/** Initialise from raw WASM bytes. Works in both Node and browser. */
export function initPptxFromBytes(bytes: Uint8Array): void {
  syncInit(pptxWasm as WasmModule, bytes, pptxState);
}
export function initDocxFromBytes(bytes: Uint8Array): void {
  syncInit(docxWasm as WasmModule, bytes, docxState);
}
export function initXlsxFromBytes(bytes: Uint8Array): void {
  syncInit(xlsxWasm as WasmModule, bytes, xlsxState);
}

function toUint8(buffer: ArrayBuffer | Uint8Array | { buffer: ArrayBuffer }): Uint8Array {
  if (buffer instanceof Uint8Array) return buffer;
  if (buffer instanceof ArrayBuffer) return new Uint8Array(buffer);
  return new Uint8Array((buffer as { buffer: ArrayBuffer }).buffer);
}

/** Convert a `.pptx` archive's bytes to GitHub-flavoured markdown. Title
 *  slides become `# heading`s, body shapes become nested bullets at the
 *  paragraph's `lvl`, tables become pipe tables, charts become summarised
 *  bullets, speaker notes and comments are collated. */
export function pptxToMarkdown(buffer: ArrayBuffer | Uint8Array | Buffer): string {
  if (!pptxState.initialized) throw new Error('pptx wasm not initialized — call initPptxFromBytes() first');
  return (pptxWasm as unknown as { pptx_to_markdown: (b: Uint8Array) => string }).pptx_to_markdown(toUint8(buffer));
}

/** Convert a `.docx` archive's bytes to GitHub-flavoured markdown. Headings
 *  come from `<w:outlineLvl>`, lists honour the abstractNum format, tables
 *  preserve vMerge continuation, footnotes/endnotes/comments are collated. */
export function docxToMarkdown(buffer: ArrayBuffer | Uint8Array | Buffer): string {
  if (!docxState.initialized) throw new Error('docx wasm not initialized — call initDocxFromBytes() first');
  return (docxWasm as unknown as { docx_to_markdown: (b: Uint8Array) => string }).docx_to_markdown(toUint8(buffer));
}

/** Convert a `.xlsx` archive's bytes to GitHub-flavoured markdown. Each
 *  sheet becomes a `## SheetName` section followed by a pipe table of its
 *  populated bbox; fully-empty middle rows are trimmed, ULP noise is
 *  masked. */
export function xlsxToMarkdown(buffer: ArrayBuffer | Uint8Array | Buffer): string {
  if (!xlsxState.initialized) throw new Error('xlsx wasm not initialized — call initXlsxFromBytes() first');
  return (xlsxWasm as unknown as { xlsx_to_markdown: (b: Uint8Array) => string }).xlsx_to_markdown(toUint8(buffer));
}
