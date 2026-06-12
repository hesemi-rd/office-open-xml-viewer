import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

// Worker mode must produce (near-)identical pixels to main mode: same
// renderer, same fonts, different thread. Slides 1 and 3 come out
// bit-identical (0.000%). Slide 2 was observed at ~0.506% in this
// configuration — sub-pixel text rasterization differing between the
// main-thread <canvas> and the worker OffscreenCanvas, confined to the
// smallest body text; the 0.7% bound leaves headroom for AA/font-hinting
// variance. Per-slide tolerances grant that slack only to the slide that
// needs it, so the bit-identical slides keep enough sensitivity that even
// a single dropped small text element (a few tenths of a percent) fails.
const SLIDES = [0, 1, 2];
const MAX_DIFF_PCT = [0.2, 0.7, 0.2];

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
