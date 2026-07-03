import { describe, it, expect, beforeAll } from 'vitest';
import { crc32 } from 'node:zlib';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { loadWasmModule, resolveWasm } from './wasm-loader.ts';

/**
 * SC18 contract: the wasm-bindgen glue for `{Pptx,Docx,Xlsx}Archive.extract_*`
 * returns an INDEPENDENT `Uint8Array` — the glue does
 * `getArrayU8FromWasm0(ptr, len).slice()` then `__wbindgen_free`s the Rust Vec,
 * so the returned array (a) no longer aliases WASM linear memory and (b) is
 * full-span over its own `ArrayBuffer` (byteOffset 0, byteLength === buffer
 * length). Those two facts are exactly what let the three `worker.ts` files
 * transfer `bytes.buffer` DIRECTLY instead of re-copying via
 * `new Uint8Array(bytes).slice().buffer`. This test pins the contract so a
 * future wasm-bindgen upgrade that returned a memory VIEW (which would make a
 * direct transfer unsafe / throw) fails here loudly.
 */

// Minimal one-entry STORED zip, same builder shape as source-buffer-image.test.
function makeZipWithEntry(name: string, data: Uint8Array): Uint8Array {
  const enc = new TextEncoder();
  const nameBytes = enc.encode(name);
  const crc = crc32(Buffer.from(data)) >>> 0;
  const size = data.length;

  const local = new Uint8Array(30 + nameBytes.length + size);
  const lv = new DataView(local.buffer);
  lv.setUint32(0, 0x04034b50, true);
  lv.setUint16(4, 20, true);
  lv.setUint16(8, 0, true); // stored
  lv.setUint32(14, crc, true);
  lv.setUint32(18, size, true);
  lv.setUint32(22, size, true);
  lv.setUint16(26, nameBytes.length, true);
  local.set(nameBytes, 30);
  local.set(data, 30 + nameBytes.length);

  const central = new Uint8Array(46 + nameBytes.length);
  const cv = new DataView(central.buffer);
  cv.setUint32(0, 0x02014b50, true);
  cv.setUint16(4, 20, true);
  cv.setUint16(6, 20, true);
  cv.setUint32(16, crc, true);
  cv.setUint32(20, size, true);
  cv.setUint32(24, size, true);
  cv.setUint16(28, nameBytes.length, true);
  central.set(nameBytes, 46);

  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, 1, true);
  ev.setUint16(10, 1, true);
  ev.setUint32(12, central.length, true);
  ev.setUint32(16, local.length, true);

  const out = new Uint8Array(local.length + central.length + eocd.length);
  out.set(local, 0);
  out.set(central, local.length);
  out.set(eocd, local.length + central.length);
  return out;
}

interface ArchiveHandle {
  extract_image(path: string): Uint8Array;
  free(): void;
}
interface PptxHandle extends ArchiveHandle {
  extract_media(path: string): Uint8Array;
}

let wasmReady = false;
beforeAll(() => {
  try {
    loadWasmModule(
      pptxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown },
      resolveWasm(import.meta.url, '../../pptx/src/wasm/pptx_parser_bg.wasm'),
    );
    loadWasmModule(
      docxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown },
      resolveWasm(import.meta.url, '../../docx/src/wasm/docx_parser_bg.wasm'),
    );
    loadWasmModule(
      xlsxWasm as unknown as { initSync: (m: WebAssembly.Module) => unknown },
      resolveWasm(import.meta.url, '../../xlsx/src/wasm/xlsx_parser_bg.wasm'),
    );
    wasmReady = true;
  } catch {
    wasmReady = false;
  }
});

/** Assert the returned array is byte-correct AND a transfer-safe, full-span,
 *  non-WASM-aliasing buffer (the SC18 precondition for a direct transfer). */
function assertIndependentFullSpan(bytes: Uint8Array, expected: Uint8Array) {
  // Byte-correct: the copy the glue made preserves the entry contents.
  expect(Array.from(bytes)).toEqual(Array.from(expected));
  // Full-span: a bare `.buffer` transfer carries exactly these bytes, nothing
  // else (no shared arena, no leading/trailing slop from a subarray view).
  expect(bytes.byteOffset).toBe(0);
  expect(bytes.byteLength).toBe(bytes.buffer.byteLength);
  // Independent of WASM linear memory: an ArrayBuffer is transferable only if it
  // is not WASM-backed. structuredClone with transfer succeeds here and detaches
  // the buffer — it would throw for a WASM-memory view.
  const buf = bytes.buffer;
  expect(() => structuredClone(buf, { transfer: [buf] })).not.toThrow();
  expect(buf.byteLength).toBe(0); // detached by the transfer → was standalone
}

describe('SC18 extract_* returns a transfer-safe buffer', () => {
  it('pptx: extract_image + extract_media are independent full-span copies', () => {
    if (!wasmReady) return;
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4, 5, 6]);
    const mp4 = new Uint8Array([0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70]);
    // Each path lives in its own one-entry zip (the builder is single-entry).
    const zip = makeZipWithEntry('ppt/media/image1.png', png);
    const Handle = (pptxWasm as unknown as { PptxArchive: new (b: Uint8Array) => PptxHandle })
      .PptxArchive;
    const ar = new Handle(zip);
    try {
      assertIndependentFullSpan(ar.extract_image('ppt/media/image1.png'), png);
    } finally {
      ar.free();
    }

    const zip2 = makeZipWithEntry('ppt/media/media2.mp4', mp4);
    const ar2 = new Handle(zip2);
    try {
      assertIndependentFullSpan(ar2.extract_media('ppt/media/media2.mp4'), mp4);
    } finally {
      ar2.free();
    }
  });

  it('docx: extract_image is an independent full-span copy', () => {
    if (!wasmReady) return;
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9, 8, 7]);
    const zip = makeZipWithEntry('word/media/image1.png', png);
    const Handle = (docxWasm as unknown as { DocxArchive: new (b: Uint8Array) => ArchiveHandle })
      .DocxArchive;
    const ar = new Handle(zip);
    try {
      assertIndependentFullSpan(ar.extract_image('word/media/image1.png'), png);
    } finally {
      ar.free();
    }
  });

  it('xlsx: extract_image is an independent full-span copy', () => {
    if (!wasmReady) return;
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 42, 43]);
    const zip = makeZipWithEntry('xl/media/image1.png', png);
    const Handle = (xlsxWasm as unknown as { XlsxArchive: new (b: Uint8Array) => ArchiveHandle })
      .XlsxArchive;
    const ar = new Handle(zip);
    try {
      assertIndependentFullSpan(ar.extract_image('xl/media/image1.png'), png);
    } finally {
      ar.free();
    }
  });
});
