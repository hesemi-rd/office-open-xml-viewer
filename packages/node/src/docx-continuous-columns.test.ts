import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasFactory,
} from './render.ts';
import { importForTests, loadSkiaForTests } from './test-imports';

// skia-canvas is a devDependency, so `pnpm install` provides it in CI as well as
// locally; the private journal samples are git-ignored (not redistributable), so
// this suite still self-skips where they are absent. Load skia through the shared
// helper: absent → skip cleanly (local), OOXML_REQUIRE_SKIA=1 (CI) → hard failure.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buf as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const RENDERER_PATH = resolve(ROOT, 'packages/docx/src/renderer.ts');
// The WASM-backed docx parser + renderer are only loaded when skia is present.
// Both statically import git-ignored WASM glue, so they need `pnpm build:wasm`
// first; under OOXML_REQUIRE_SKIA=1 a failure to load is a hard error.
const docxMod = skia ? await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)') : null;
const rendererMod = skia
  ? await importForTests(() => import(RENDERER_PATH), 'packages/docx/src/renderer.ts')
  : null;

const samplePath = (n: number) =>
  resolve(ROOT, `packages/docx/public/private/sample-${n}.docx`);
const haveSamples =
  existsSync(samplePath(5)) && existsSync(samplePath(12)) && existsSync(samplePath(13));

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

    // 5 pages, matching Word. The intro 2-col section opens with a "continuous"
    // section break, so it stays on the title page (§17.6.22: the break is
    // governed by the upcoming section's start type, not the title section's
    // nextPage). Restored after the sample-5 cover overprint was fixed at its
    // real root — a PageBreak after the "Cover Pages" building block (§17.5.2) —
    // instead of forcing every nextPage→continuous boundary to break a page.
    it('sample-13 paginates to 5 pages (Word ground truth)', () => {
      expect(paginate(13).length).toBe(5);
    });

    // sample-5 (夢十夜): the cover is a "Cover Pages" building block (§17.5.2)
    // whose text flow is empty — the page is filled by page-anchored cover
    // graphics. Word places it on its own page and starts the novel body on
    // page 2, even though the body section opens with a "continuous" break. The
    // parser emits a PageBreak after the cover content so the cover stands alone:
    // 7 pages. Were the cover detection to fail, the continuous body would flow
    // up onto page 1 and the document would collapse to 6 pages.
    it('sample-5 cover page stands alone — 7 pages (Word ground truth)', () => {
      expect(paginate(5).length).toBe(7);
    });
  },
);
