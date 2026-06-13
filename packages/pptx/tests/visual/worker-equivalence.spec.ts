import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

// Worker mode must produce identical pixels to main mode: same renderer, same
// fonts (the worker loads the same web fonts into its OffscreenCanvas font set),
// different thread. All three slides come out bit-identical (0.000%); the tiny
// uniform tolerance only absorbs rare AA/hinting noise and still fails on a
// dropped element or a font that didn't load in the worker (which would diverge
// by whole tenths of a percent — that was the symptom before the preload fix).
const SLIDES = [0, 1, 2];
const MAX_DIFF_PCT = [0.1, 0.1, 0.1];

for (const slide of SLIDES) {
  test(`worker mode matches main mode › demo/sample-1 slide ${slide + 1}`, async ({ page }) => {
    await page.goto(`/tests/visual/worker-fixture.html?pptx=demo/sample-1&slide=${slide}`);
    // Two full loads (main + worker) plus worker spin-up per test — twice the
    // single-render budget visual.spec.ts uses.
    await page.waitForFunction(
      () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
      { timeout: 60_000 },
    );
    const status = await page.evaluate(() => document.body.dataset.status);
    if (status === 'error') {
      throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
    }

    const [mainUrl, workerUrl] = await page.evaluate(() => [
      (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
      (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    ]);
    const a = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
    const b = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
    // A zero-size canvas means a silently failed render; fail with a readable
    // assertion instead of the NaN the diff percentage would produce.
    expect(a.width).toBeGreaterThan(0);
    expect(a.height).toBeGreaterThan(0);
    expect(b.width).toBe(a.width);
    expect(b.height).toBe(a.height);

    const diff = pixelmatch(a.data, b.data, undefined, a.width, a.height, { threshold: 0.1 });
    const pct = (diff / (a.width * a.height)) * 100;
    console.log(`  slide ${slide + 1}: worker-vs-main diff ${pct.toFixed(3)}%`);
    expect(pct).toBeLessThanOrEqual(MAX_DIFF_PCT[slide]);
  });
}
