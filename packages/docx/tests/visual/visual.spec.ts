import { test } from '@playwright/test';
import { mkdirSync, existsSync, readFileSync, writeFileSync } from 'fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

const DOCX_FILES: { name: string; pageCount: number; width: number }[] = [
  { name: 'private/sample-1', pageCount: 1, width: 612 },
  { name: 'private/sample-2', pageCount: 1, width: 595 },
  { name: 'private/sample-3', pageCount: 3, width: 595 },
  { name: 'private/sample-4', pageCount: 1, width: 595 },
  { name: 'private/sample-5', pageCount: 7, width: 595 },
  // CH13 chart coverage: sample-24 p.3 carries a stockChart (hi-lo-close);
  // sample-25 is a pie3DChart. References are private (gitignored) and generated
  // locally with UPDATE_REFS=1 — they are never committed.
  { name: 'private/sample-24', pageCount: 3, width: 595 },
  { name: 'private/sample-25', pageCount: 1, width: 595 },
  { name: 'demo/sample-1', pageCount: 6, width: 595 },
];

const PIXEL_THRESHOLD = 0.20;
const FAIL_ABOVE_PCT = 20;
const REGRESSION_PCT = 0.5;
// Fidelity-score ratchet: fail if a page's match-% vs its reference PNG drops
// more than this below the committed score. Catches a renderer change that
// quietly worsens fidelity against the Word ground truth even while staying
// under the coarse FAIL_ABOVE_PCT ceiling.
const RATCHET_DROP_PCT = 0.5;

// UPDATE_REFS=1 pnpm vrt → adopt the current canvas output as the new reference.
const UPDATE_REFS = process.env.UPDATE_REFS === '1';
// UPDATE_SCORES=1 pnpm vrt → record the current fidelity match-% into
// references/<name>/scores.json WITHOUT touching the reference PNGs. This is how
// the committed demo scores are (re)generated; it never rewrites ground truth.
const UPDATE_SCORES = process.env.UPDATE_SCORES === '1';
const SNAPSHOT = process.env.VRT_SNAPSHOT === '1';
const RUN_MODE = process.env.VRT_MODE === 'regression' ? 'regression' : 'fidelity';

// Per-sample fidelity scores live next to the reference PNGs
// (references/<name>/scores.json), so they inherit the exact same commit policy:
// demo scores are tracked, private scores are gitignored. Keyed by item id
// (e.g. "page-3") → match-% (2 dp). Read-modify-write is safe because the VRT
// config runs sequentially (fullyParallel: false).
function scoresPathFor(name: string): string {
  return `tests/visual/references/${name}/scores.json`;
}
function readScores(name: string): Record<string, number> {
  const p = scoresPathFor(name);
  if (!existsSync(p)) return {};
  try {
    return JSON.parse(readFileSync(p, 'utf8')) as Record<string, number>;
  } catch {
    return {};
  }
}
function writeScore(name: string, key: string, matchPct: number): void {
  const scores = readScores(name);
  scores[key] = Math.round(matchPct * 100) / 100;
  mkdirSync(`tests/visual/references/${name}`, { recursive: true });
  const ordered = Object.fromEntries(Object.entries(scores).sort(([a], [b]) => a.localeCompare(b)));
  writeFileSync(scoresPathFor(name), JSON.stringify(ordered, null, 2) + '\n');
}

test.describe('docx visual regression', () => {
  for (const { name, pageCount, width } of DOCX_FILES) {
    for (let i = 0; i < pageCount; i++) {
      const pageNum = i + 1;

      test(`${name} › page ${pageNum}`, async ({ page }) => {
        await page.goto(
          `/tests/visual/fixture.html?file=${name}.docx&page=${i}&width=${width}`
        );

        await page.waitForFunction(
          () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
          { timeout: 30_000 }
        );

        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error on ${name} page ${pageNum}: ${msg}`);
        }

        await page.waitForTimeout(200);

        const dataUrl = await page.evaluate(() => {
          const canvas = document.querySelector('canvas') as HTMLCanvasElement;
          return canvas ? canvas.toDataURL('image/png') : null;
        });
        if (!dataUrl) throw new Error(`No canvas on ${name} page ${pageNum}`);
        const actualBuf = Buffer.from(dataUrl.split(',')[1], 'base64');

        mkdirSync(`tests/visual/screenshots/${name}`, { recursive: true });
        writeFileSync(`tests/visual/screenshots/${name}/page-${pageNum}.png`, actualBuf);

        if (UPDATE_REFS) {
          mkdirSync(`tests/visual/references/${name}`, { recursive: true });
          writeFileSync(`tests/visual/references/${name}/page-${pageNum}.png`, actualBuf);
          console.log(`  ${name} page ${pageNum}: reference updated`);
          return;
        }
        if (SNAPSHOT) {
          mkdirSync(`tests/visual/baseline/${name}`, { recursive: true });
          writeFileSync(`tests/visual/baseline/${name}/page-${pageNum}.png`, actualBuf);
          console.log(`  ${name} page ${pageNum}: baseline captured`);
          return;
        }

        const targetRoot = RUN_MODE === 'regression' ? 'baseline' : 'references';
        const refPath = `tests/visual/${targetRoot}/${name}/page-${pageNum}.png`;
        if (!existsSync(refPath)) {
          test.skip(true, `no ${targetRoot} image for ${name} page ${pageNum}`);
        }
        const refBuf = readFileSync(refPath);
        const refPng    = PNG.sync.read(refBuf);
        const actualPng = PNG.sync.read(actualBuf);

        const { width: refW, height: refH } = refPng;

        if (actualPng.width !== refW || actualPng.height !== refH) {
          console.warn(
            `  ${name} page ${pageNum}: size mismatch ` +
            `actual=${actualPng.width}×${actualPng.height} ` +
            `ref=${refW}×${refH}`
          );
        }

        const w = Math.min(actualPng.width, refW);
        const h = Math.min(actualPng.height, refH);

        // Pad both images to same size so pixelmatch doesn't throw
        const pad = (png: ReturnType<typeof PNG.sync.read>, tw: number, th: number) => {
          if (png.width === tw && png.height === th) return png;
          const out = new PNG({ width: tw, height: th });
          out.data.fill(255);
          for (let y = 0; y < Math.min(png.height, th); y++) {
            for (let x = 0; x < Math.min(png.width, tw); x++) {
              const src = (y * png.width + x) * 4;
              const dst = (y * tw + x) * 4;
              out.data[dst]     = png.data[src];
              out.data[dst + 1] = png.data[src + 1];
              out.data[dst + 2] = png.data[src + 2];
              out.data[dst + 3] = png.data[src + 3];
            }
          }
          return out;
        };
        const refPadded    = pad(refPng,    w, h);
        const actualPadded = pad(actualPng, w, h);

        const diff = new PNG({ width: w, height: h });
        const diffPixels = pixelmatch(
          refPadded.data, actualPadded.data, diff.data, w, h,
          { threshold: PIXEL_THRESHOLD, includeAA: true }
        );
        mkdirSync(`tests/visual/diffs/${name}`, { recursive: true });
        writeFileSync(`tests/visual/diffs/${name}/page-${pageNum}.png`, PNG.sync.write(diff));

        const totalPx = w * h;
        const diffPct = (diffPixels / totalPx) * 100;
        const matchPct = 100 - diffPct;

        console.log(
          `  ${name} page ${pageNum}: ` +
          `match=${matchPct.toFixed(1)}%  diff=${diffPct.toFixed(1)}%  ` +
          `(${diffPixels.toLocaleString()} / ${totalPx.toLocaleString()} px)`
        );

        const limit = RUN_MODE === 'regression' ? REGRESSION_PCT : FAIL_ABOVE_PCT;
        if (diffPct > limit) {
          throw new Error(
            `${name} page ${pageNum} pixel diff ${diffPct.toFixed(1)}% exceeds ${limit}% in ${RUN_MODE} mode`
          );
        }

        // Fidelity-score ratchet (fidelity mode only; the regression mode above
        // already gates against the captured baseline). UPDATE_SCORES rewrites
        // the stored score; otherwise a committed score is a floor.
        if (RUN_MODE === 'fidelity') {
          const key = `page-${pageNum}`;
          if (UPDATE_SCORES) {
            writeScore(name, key, matchPct);
          } else {
            const prior = readScores(name)[key];
            if (prior !== undefined && matchPct < prior - RATCHET_DROP_PCT) {
              throw new Error(
                `${name} ${key} fidelity regressed: match ${matchPct.toFixed(2)}% ` +
                `is >${RATCHET_DROP_PCT}pt below the recorded ${prior.toFixed(2)}%`
              );
            }
          }
        }
      });
    }
  }
});
