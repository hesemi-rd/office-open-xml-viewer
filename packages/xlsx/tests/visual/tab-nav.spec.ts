import { test, expect, type Locator } from '@playwright/test';

test.describe('xlsx sheet tab navigation buttons', () => {
  test('prev/next scroll the tab strip and disable at the ends', async ({ page }) => {
    await page.goto('/tests/visual/viewer-fixture.html');
    await expect(page.locator('body')).toHaveAttribute('data-status', 'ready');

    const prev = page.locator('[data-xlsx-tab-nav="prev"]');
    const next = page.locator('[data-xlsx-tab-nav="next"]');
    await expect(prev).toBeVisible();
    await expect(next).toBeVisible();

    const scrollLeft = () =>
      page.evaluate(() => {
        const strip = document.querySelector('.xlsx-tab-strip');
        return strip ? (strip as HTMLElement).scrollLeft : -1;
      });

    // Disabled is encoded as pointer-events:none on the button's inline style.
    const isDisabled = (loc: Locator) =>
      loc.evaluate((el) => (el as HTMLElement).style.pointerEvents === 'none');

    // Initial state: scrolled to start, prev disabled, next enabled (5 long
    // tabs overflow the 360px host).
    await expect.poll(scrollLeft).toBe(0);
    await expect.poll(() => isDisabled(prev)).toBe(true);
    await expect.poll(() => isDisabled(next)).toBe(false);

    // One next click must increase scrollLeft.
    const before = await scrollLeft();
    await next.click();
    await expect.poll(scrollLeft).toBeGreaterThan(before);

    // Click next until it disables (reached the end).
    for (let i = 0; i < 12; i++) {
      if (await isDisabled(next)) break;
      await next.click();
    }
    await expect.poll(() => isDisabled(next)).toBe(true);
    await expect.poll(() => isDisabled(prev)).toBe(false);

    // Click prev back to the start.
    for (let i = 0; i < 12; i++) {
      if (await isDisabled(prev)) break;
      await prev.click();
    }
    await expect.poll(() => isDisabled(prev)).toBe(true);
    await expect.poll(scrollLeft).toBe(0);
  });
});
