import { test, expect } from '@playwright/test';

/**
 * IX6 functional worker-equivalence: the text-selection overlay and findText
 * must produce the SAME result in `mode: 'worker'` as in `mode: 'main'`. The
 * run geometry is collected off-thread in worker mode and shipped back beside
 * the bitmap, so the selection spans (text + rounded shape/in-shape box) and the
 * findText match set must be byte-for-byte identical to main mode. This is the
 * functional twin of the pixel worker-equivalence spec: it fails if a run is
 * dropped on the wire or findText silently returns `[]` in worker mode (the
 * pre-IX6 behaviour).
 */
test('worker mode selection overlay + findText match main mode › demo/sample-1', async ({ page }) => {
  await page.goto('/tests/visual/worker-selection-fixture.html?pptx=demo/sample-1&q=a');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }

  const data = await page.evaluate(() => ({
    mainOverlay: document.body.dataset.mainOverlay ?? '[]',
    workerOverlay: document.body.dataset.workerOverlay ?? '[]',
    mainMatches: document.body.dataset.mainMatches ?? '[]',
    workerMatches: document.body.dataset.workerMatches ?? '[]',
  }));

  const mainOverlay = JSON.parse(data.mainOverlay) as unknown[];
  const workerOverlay = JSON.parse(data.workerOverlay) as unknown[];
  const mainMatches = JSON.parse(data.mainMatches) as unknown[];
  const workerMatches = JSON.parse(data.workerMatches) as unknown[];

  // The overlay must be non-empty (a real slide with selectable text), or the
  // equivalence below would be a vacuous [] === [].
  expect(mainOverlay.length).toBeGreaterThan(0);
  // The selection overlay is identical run-for-run between modes.
  expect(workerOverlay).toEqual(mainOverlay);

  // findText found real matches in main mode…
  expect(mainMatches.length).toBeGreaterThan(0);
  // …and worker mode returns the exact same match set (this is the guard the
  // pre-IX6 code broke: worker findText returned [] + a console.warn).
  expect(workerMatches).toEqual(mainMatches);
});
