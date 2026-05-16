#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { resolve, dirname, extname } from 'node:path';
import { fileURLToPath } from 'node:url';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', short: 'o' },
    help: { type: 'boolean', short: 'h' },
  },
});

if (values.help || positionals.length === 0) {
  console.log(`ooxml-md — convert .pptx / .docx / .xlsx to GitHub-flavoured markdown

Usage:
  ooxml-md <file>              # writes to stdout
  ooxml-md <file> -o out.md    # writes to file
`);
  process.exit(values.help ? 0 : 1);
}

const filePath = resolve(positionals[0]);
const ext = extname(filePath).toLowerCase();
const here = dirname(fileURLToPath(import.meta.url));

const {
  pptxToMarkdown,
  docxToMarkdown,
  xlsxToMarkdown,
  initPptxFromBytes,
  initDocxFromBytes,
  initXlsxFromBytes,
} = await import('../src/index.ts').catch(() => import('../src/index.js'));

const buf = readFileSync(filePath);
let md;
if (ext === '.pptx') {
  const wasm = readFileSync(resolve(here, '../../pptx/src/wasm/pptx_parser_bg.wasm'));
  initPptxFromBytes(wasm);
  md = pptxToMarkdown(buf);
} else if (ext === '.docx') {
  const wasm = readFileSync(resolve(here, '../../docx/src/wasm/docx_parser_bg.wasm'));
  initDocxFromBytes(wasm);
  md = docxToMarkdown(buf);
} else if (ext === '.xlsx') {
  const wasm = readFileSync(resolve(here, '../../xlsx/src/wasm/xlsx_parser_bg.wasm'));
  initXlsxFromBytes(wasm);
  md = xlsxToMarkdown(buf);
} else {
  console.error(`Unsupported extension: ${ext}. Expected .pptx / .docx / .xlsx`);
  process.exit(2);
}

if (values.out) {
  writeFileSync(resolve(values.out), md);
  console.error(`Wrote ${md.length} bytes to ${values.out}`);
} else {
  process.stdout.write(md);
}
