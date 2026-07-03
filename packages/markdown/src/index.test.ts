import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import {
  docxToMarkdown,
  pptxToMarkdown,
  xlsxToMarkdown,
  initDocxFromBytes,
  initPptxFromBytes,
  initXlsxFromBytes,
} from './index.ts';

/**
 * End-to-end unit coverage for the markdown conversion functions: init each
 * format's WASM from disk (the same `./wasm-binary` export the CLI resolves),
 * then convert the committed demo samples and assert the projection carries the
 * expected top-level structure.
 *
 * The parser WASM is git-ignored build output; CI runs `pnpm build:wasm` before
 * `pnpm test`, so it is present there, but a local run before a wasm build would
 * lack it. Rather than hard-fail, we gate the suite so a missing artifact SKIPS
 * (never a silent zero-assertion pass) — mirroring the node package's
 * `archive-extract-transfer.test.ts`.
 */

const require = createRequire(import.meta.url);
const root = new URL('../../..', import.meta.url); // repo root from packages/markdown/src

function tryRead(pkgWasm: string): Uint8Array | null {
  try {
    return new Uint8Array(readFileSync(require.resolve(pkgWasm)));
  } catch {
    return null;
  }
}
function trySample(rel: string): Buffer | null {
  try {
    return readFileSync(new URL(rel, root));
  } catch {
    return null;
  }
}

const docxWasm = tryRead('@silurus/ooxml-docx/wasm-binary');
const pptxWasm = tryRead('@silurus/ooxml-pptx/wasm-binary');
const xlsxWasm = tryRead('@silurus/ooxml-xlsx/wasm-binary');
const docxSample = trySample('packages/docx/public/demo/sample-1.docx');
const pptxSample = trySample('packages/pptx/public/demo/sample-1.pptx');
const xlsxSample = trySample('packages/xlsx/public/demo/sample-1.xlsx');

const docxReady = !!docxWasm && !!docxSample;
const pptxReady = !!pptxWasm && !!pptxSample;
const xlsxReady = !!xlsxWasm && !!xlsxSample;

describe('markdown conversion', () => {
  it.skipIf(!docxReady)('converts a .docx to markdown', () => {
    initDocxFromBytes(docxWasm as Uint8Array);
    const md = docxToMarkdown(docxSample as Buffer);
    expect(md.length).toBeGreaterThan(0);
    // The demo sample's masthead survives the projection.
    expect(md).toContain('CANOPY');
  });

  it.skipIf(!pptxReady)('converts a .pptx to markdown', () => {
    initPptxFromBytes(pptxWasm as Uint8Array);
    const md = pptxToMarkdown(pptxSample as Buffer);
    expect(md.length).toBeGreaterThan(0);
    // Slide titles become `#` headings.
    expect(md).toMatch(/^#\s/m);
  });

  it.skipIf(!xlsxReady)('converts a .xlsx to markdown', () => {
    initXlsxFromBytes(xlsxWasm as Uint8Array);
    const md = xlsxToMarkdown(xlsxSample as Buffer);
    expect(md.length).toBeGreaterThan(0);
    // Each sheet becomes a `## SheetName` section with a pipe table.
    expect(md).toMatch(/^##\s/m);
    expect(md).toContain('|');
  });
});
