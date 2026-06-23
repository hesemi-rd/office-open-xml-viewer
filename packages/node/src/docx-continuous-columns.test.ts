import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
} from './render.ts';

// skia-canvas is a local-only peer dep (CI omits it); the private journal
// samples are git-ignored (not redistributable). Skip cleanly when either is
// absent — this suite is a local ground-truth gate, like the VRT specs.
const skia = await import('skia-canvas').catch(() => null);
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buf as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const docxMod = skia ? await import('./docx.ts').catch(() => null) : null;
const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const RENDERER_PATH = resolve(ROOT, 'packages/docx/src/renderer.ts');
const rendererMod = skia ? await import(RENDERER_PATH).catch(() => null) : null;

const samplePath = (n: number) =>
  resolve(ROOT, `packages/docx/public/private/sample-${n}.docx`);
const haveSamples = existsSync(samplePath(12)) && existsSync(samplePath(13));

// ECMA-376 §17.6.4 newspaper columns + §17.18.79 "continuous" section marks.
// Both journal templates flow their body through `continuous` section breaks
// that flip the column count (1 ⇄ 2) mid-page. The paginator must place each
// multi-column region's later columns at the REGION top (where the section
// began on the page), not the page content top — otherwise the second column
// overprints the preceding single-column content and the page absorbs too much
// content (sample-12 collapsed 3 Word pages into 2). Ground truth = the Word
// PDF page counts next to each .docx.
describe.skipIf(!skia || !docxMod || !rendererMod || !haveSamples)(
  'continuous column-count section breaks (sample-12/13)',
  () => {
    let restore: Array<() => void> = [];
    const paginate = (n: number) => {
      restore = [installOffscreenCanvasShim(factory), installImageBitmapShim(factory)];
      try {
        const { parseDocx } = docxMod!;
        const { paginateDocument } = rendererMod as {
          paginateDocument: (doc: unknown) => unknown[][];
        };
        const doc = parseDocx(readFileSync(samplePath(n)));
        return paginateDocument(doc);
      } finally {
        restore.forEach((r) => r());
      }
    };

    // Tier 1 (column-region top tracking): the second column of a continuous
    // mid-page multi-column section starts at the region top, not the page top —
    // so the overprint is gone and sample-12 flows across its 3 Word pages.
    it('sample-12 paginates to 3 pages (Word ground truth)', () => {
      expect(paginate(12).length).toBe(3);
    });

    // Now 6 pages (was 7 before the region-top fix, 5 in Word). The overprint is
    // gone and the §17.6.22 boundary fix keeps the 2-col body on the title page,
    // but one extra page remains in the figure/table-dense middle section. Word
    // does NOT balance these columns (its last 2-col page fills the left column
    // only), so the residual is a figure/table-sizing + greedy-fill fidelity gap,
    // not balancing. Marked `fails` as a tracking gate against Word's 5 pages.
    it.fails('sample-13 paginates to 5 pages — residual figure-region fidelity gap', () => {
      expect(paginate(13).length).toBe(5);
    });
  },
);
