import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

// Worker mode must produce (near-)identical pixels to main mode: same
// renderer, same per-sheet parse, same fonts, different thread. The only
// expected drift is sub-pixel text rasterization between the main-thread
// <canvas> and the worker OffscreenCanvas. Per-sheet tolerances grant slack
// only to a sheet that needs it, so the (near-)bit-identical sheets keep
// enough sensitivity that even a single dropped cell fails the diff.
const SHEETS = [0, 1];
const MAX_DIFF_PCT = [0.5, 0.5];

for (const sheetIndex of SHEETS) {
  test(`worker mode matches main mode › demo/sample-1 sheet ${sheetIndex + 1}`, async ({ page }) => {
    await page.goto(`/tests/visual/worker-fixture.html?xlsx=demo/sample-1&sheet=${sheetIndex}`);
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

    // The fixture also drives a worker-mode getWorksheet() round-trip, which
    // posts a `parseSheet` message the render worker must handle. If that arm
    // regresses the promise never settles, the fixture stalls before 'ready'
    // and the waitForFunction above fails loudly. Assert the probe resolved
    // with a real worksheet so a silently-empty result can't slip through.
    const worksheetOk = await page.evaluate(() => document.body.dataset.worksheetOk);
    expect(worksheetOk).toBe('true');

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
    console.log(`  sheet ${sheetIndex + 1}: worker-vs-main diff ${pct.toFixed(3)}%`);
    expect(pct).toBeLessThanOrEqual(MAX_DIFF_PCT[sheetIndex]);
  });
}
