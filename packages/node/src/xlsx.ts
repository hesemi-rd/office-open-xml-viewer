import type { ParsedWorkbook, Worksheet } from '@silurus/ooxml-xlsx';
// The package typecheck alias maps '@silurus/ooxml-xlsx' to types.ts (types
// only), so the resolver value is imported from source directly — mirroring the
// relative WASM import below. (index.ts re-exports it for external consumers.)
import { resolveSharedStrings } from '../../xlsx/src/shared-strings.ts';
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

let initialized = false;

function ensureInit(): void {
  if (initialized) return;
  const wasmPath = resolveWasm(import.meta.url, '../../xlsx/src/wasm/xlsx_parser_bg.wasm');
  loadWasmModule(xlsxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown }, wasmPath);
  initialized = true;
}

/** Parse the workbook index (sheet list + styles + shared strings) from a
 *  `.xlsx` archive. Individual sheet cell data is parsed lazily via
 *  {@link parseSheet}. */
export function parseXlsx(buffer: ArrayBuffer | Uint8Array | Buffer): ParsedWorkbook {
  ensureInit();
  const bytes = toUint8(buffer);
  // `parse_xlsx` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once. Matches the browser main-thread receiver.
  const json = (xlsxWasm as unknown as { parse_xlsx: (b: Uint8Array) => Uint8Array }).parse_xlsx(
    bytes,
  );
  return JSON.parse(new TextDecoder().decode(json)) as ParsedWorkbook;
}

/** Parse a single sheet's cell model without resolving shared-string cells
 *  (they stay `{type:'shared',si}`). Internal — callers resolve against the
 *  workbook's sharedStrings table. */
function parseSheetRaw(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  sheetIndex: number,
  sheetName: string,
): Worksheet {
  ensureInit();
  const bytes = toUint8(buffer);
  // `parse_sheet` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>); decode +
  // parse once.
  const json = (xlsxWasm as unknown as {
    parse_sheet: (b: Uint8Array, idx: number, name: string) => Uint8Array;
  }).parse_sheet(bytes, sheetIndex, sheetName);
  return JSON.parse(new TextDecoder().decode(json)) as Worksheet;
}

/** Parse a single sheet's cell data and layout. The browser path does this
 *  on demand from a Web Worker; in Node we just call the WASM export
 *  synchronously. Shared-string cells are resolved to concrete text against the
 *  workbook table (matching the browser `XlsxWorkbook` path), so callers reading
 *  `cell.value` always get `{type:'text',...}` rather than a `{type:'shared'}`
 *  reference. */
export function parseSheet(
  buffer: ArrayBuffer | Uint8Array | Buffer,
  sheetIndex: number,
  sheetName: string,
): Worksheet {
  const ws = parseSheetRaw(buffer, sheetIndex, sheetName);
  return resolveSharedStrings(ws, parseXlsx(buffer).sharedStrings);
}

/** Eagerly parse every sheet referenced by the workbook. Useful for batch
 *  jobs (diffing two workbooks, dumping to markdown) where you want the
 *  whole model in one go. */
export function parseXlsxAllSheets(
  buffer: ArrayBuffer | Uint8Array | Buffer,
): { workbook: ParsedWorkbook['workbook']; worksheets: Record<string, Worksheet> } {
  const parsed = parseXlsx(buffer);
  const worksheets: Record<string, Worksheet> = {};
  for (let i = 0; i < parsed.workbook.sheets.length; i++) {
    const meta = parsed.workbook.sheets[i];
    // Resolve against the already-parsed table (avoids re-parsing the workbook
    // index once per sheet, which routing through `parseSheet` would do).
    worksheets[meta.name] = resolveSharedStrings(
      parseSheetRaw(buffer, i, meta.name),
      parsed.sharedStrings,
    );
  }
  return { workbook: parsed.workbook, worksheets };
}

function toUint8(buffer: ArrayBuffer | Uint8Array | Buffer): Uint8Array {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer as ArrayBuffer);
}
