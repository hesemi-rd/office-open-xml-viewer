#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { resolve } from 'node:path';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', short: 'o' },
    slide: { type: 'string', short: 's', default: '0' },
    page: { type: 'string', short: 'p', default: '0' },
    width: { type: 'string', short: 'w', default: '960' },
    help: { type: 'boolean', short: 'h' },
  },
});

if (values.help || positionals.length === 0) {
  console.log(`ooxml-thumbnail — generate a PNG thumbnail of a pptx / docx slide / page

Usage:
  ooxml-thumbnail <file.pptx|file.docx> [--out thumb.png] [--slide N] [--page N] [--width 960]

Requires the optional peer dependency 'skia-canvas' (\`npm i skia-canvas\`).
XLSX thumbnail support is not yet implemented in the Node path.

Examples:
  ooxml-thumbnail deck.pptx --slide 0 --out s0.png
  ooxml-thumbnail report.docx --page 2 --out p2.png
`);
  process.exit(values.help ? 0 : 1);
}

const filePath = resolve(positionals[0]);
const outPath = values.out ? resolve(values.out) : filePath.replace(/\.(pptx|docx)$/i, '.png');

const buffer = readFileSync(filePath);

let skia;
try {
  skia = await import('skia-canvas');
} catch (err) {
  console.error('skia-canvas is required for thumbnail generation. Install it with:');
  console.error('  npm install skia-canvas');
  console.error('');
  console.error('Original error:', err.message);
  process.exit(1);
}

const { Canvas, loadImage, FontLibrary } = skia;

// Wire fonts: pick up system fonts so generic text renders.
// (Users who need theme fonts should call FontLibrary.use(...) themselves
// before invoking the CLI, e.g. from a wrapper script.)
void FontLibrary;

const { parsePptx, parseDocx, renderSlideNode, installImageBitmapShim } =
  await import('../src/index.ts').catch(() => import('../src/index.js'));

const factory = {
  createCanvas: (w, h) => new Canvas(w, h),
  loadImage: (buf) => loadImage(buf),
};

installImageBitmapShim(factory);

const slideIdx = Number.parseInt(values.slide, 10);
const pageIdx = Number.parseInt(values.page, 10);
const width = Number.parseInt(values.width, 10);

if (filePath.toLowerCase().endsWith('.pptx')) {
  const presentation = parsePptx(buffer);
  const slideWidth = presentation.slideWidth;
  const slideHeight = presentation.slideHeight;
  const dpr = 2;
  const cssH = Math.round(slideHeight * (width / slideWidth));
  const canvas = factory.createCanvas(width * dpr, cssH * dpr);
  await renderSlideNode(canvas, presentation, slideIdx, { width, dpr, factory });
  const png = await canvas.toBuffer('png');
  writeFileSync(outPath, png);
  console.log(`Wrote ${outPath} (slide ${slideIdx + 1} / ${presentation.slides.length})`);
} else if (filePath.toLowerCase().endsWith('.docx')) {
  void pageIdx;
  void parseDocx;
  console.error('DOCX Node rendering is not yet implemented — only PPTX is supported by the v0 thumbnail CLI.');
  process.exit(2);
} else {
  console.error(`Unsupported extension. Expected .pptx or .docx; got ${filePath}`);
  process.exit(2);
}
