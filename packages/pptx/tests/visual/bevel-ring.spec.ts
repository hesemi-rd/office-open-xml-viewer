import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

/**
 * Acceptance test for the slide-6 sp3d bevel ring (private/sample-11.pptx, the
 * tilted-ellipse photo with a 15 pt a:ln beige border, bevelT hardEdge,
 * extrusionH and outerShdw under a perspectiveFront camera).
 *
 * History: across sessions the ellipse rim regressed into a "white band" — a
 * pure-white stripe sitting in the outer half of the ring, making the
 * silhouette read as cut off against the white slide. The merged bevel
 * band-geometry / lip-azimuth PRs cured it; this spec LOCKS the cure in.
 *
 * Ground truth (sample-11.pdf p6): the ring is a CONTINUOUS pale-beige band on
 * every apex — a soft outer shadow outside it, then opaque beige, then the
 * photo, with NO white interruption and no saturated-white lit lip.
 *
 * The .pptx is private (never committed). Like the other private-sample specs
 * this one SKIPS when the fixture is absent so a clean checkout stays green.
 * It asserts pixel structure from the live canvas, not a reference image.
 */
test.describe('pptx slide-6 sp3d bevel ring', () => {
  const SAMPLE = 'private/sample-11';
  const SLIDE_INDEX = 5; // PDF p6, 0-based

  const fixturePath = join(
    dirname(fileURLToPath(import.meta.url)),
    '..',
    '..',
    'public',
    `${SAMPLE}.pptx`,
  );

  test('the ring is a continuous beige band on every apex — no white cut, no blown lit lip', async ({
    page,
  }, testInfo) => {
    testInfo.skip(!existsSync(fixturePath), `${SAMPLE}.pptx not present (private sample)`);

    await page.goto(`/tests/visual/fixture.html?pptx=${SAMPLE}&slide=${SLIDE_INDEX}`);
    await page.waitForFunction(
      () =>
        document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
      { timeout: 30_000 },
    );
    const status = await page.evaluate(() => document.body.dataset.status);
    if (status === 'error') {
      const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
      throw new Error(`Fixture error: ${msg}`);
    }
    await page.waitForTimeout(200);

    // Sample the four apex cross-sections off the live canvas and classify each
    // ray's pixels into page-white / beige-rim / photo. The acceptance
    // contract: walking from the page into the photo we cross a CONTINUOUS
    // beige run wide enough to read as the ring with ZERO white pixels in the
    // rim region, and the rim never saturates (luma < 250).
    //
    // The ring is the 15 pt = 190500 EMU centre-aligned a:ln border (the bevel
    // lip sits under it). After antialiasing toward the photo and the outer
    // shadow the PURE-beige run measures ~11 device px at deviceScaleFactor 1
    // (cw 1280) and ~40 px at dpr 2; sample-11.pdf p6 rasterised to matching
    // width shows the same band. The floor is DPR-aware (fraction of canvas
    // width) with an 8 px absolute minimum — low enough to tolerate the AA at
    // these narrow apices, high enough to fail hard if the ring collapses to a
    // thin line or is split by a white stripe (the regression we guard).
    const result = await page.evaluate(() => {
      const canvas = document.querySelector('canvas') as HTMLCanvasElement;
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
      const cw = canvas.width;
      const ch = canvas.height;
      const minRim = Math.max(8, Math.round(cw * 0.006));
      const D = ctx.getImageData(0, 0, cw, ch).data;
      const at = (x: number, y: number): [number, number, number] => {
        const o = (Math.round(y) * cw + Math.round(x)) * 4;
        return [D[o], D[o + 1], D[o + 2]];
      };
      const isWhite = (r: number, g: number, b: number) => r >= 250 && g >= 250 && b >= 250;
      const isPhoto = (r: number, g: number, b: number) => b > 95 && b > r + 25 && b > g + 15;
      const isBeige = (r: number, g: number, b: number) =>
        r > 150 && g > 150 && b > 130 && Math.abs(r - g) < 30 && r > b && !isWhite(r, g, b);

      function analyze(name: string, x0: number, y0: number, dx: number, dy: number) {
        let x = x0;
        let y = y0;
        // Skip leading page white.
        while (x >= 0 && y >= 0 && x < cw && y < ch) {
          const [r, g, b] = at(x, y);
          if (!isWhite(r, g, b)) break;
          x += dx;
          y += dy;
        }
        // March through the rim until the photo, tracking white intrusions and
        // the longest beige run + its brightest pixel.
        let whiteInRim = 0;
        let beigeRun = 0;
        let bestBeigeRun = 0;
        let beigeMax = 0;
        for (let i = 0; i < Math.max(cw, ch); i++) {
          if (x < 0 || y < 0 || x >= cw || y >= ch) break;
          const [r, g, b] = at(x, y);
          if (isPhoto(r, g, b)) break;
          if (isWhite(r, g, b)) whiteInRim++;
          if (isBeige(r, g, b)) {
            beigeRun++;
            bestBeigeRun = Math.max(bestBeigeRun, beigeRun);
            beigeMax = Math.max(beigeMax, 0.299 * r + 0.587 * g + 0.114 * b);
          } else {
            beigeRun = 0;
          }
          x += dx;
          y += dy;
        }
        return { name, whiteInRim, bestBeigeRun, beigeMax: Math.round(beigeMax) };
      }

      const cx = Math.round(cw / 2);
      const cy = Math.round(ch / 2);
      return {
        minRim,
        cw,
        sections: [
          analyze('TOP', cx, 0, 0, 1),
          analyze('BOTTOM', cx, ch - 1, 0, -1),
          analyze('LEFT', 0, cy, 1, 0),
          analyze('RIGHT', cw - 1, cy, -1, 0),
        ],
      };
    });

    console.log(
      `slide-6 ring (canvas ${result.cw}px, minRim ${result.minRim}px):`,
      JSON.stringify(result.sections),
    );
    for (const r of result.sections) {
      // 1) No white band inside the rim (this is the regression we guard).
      expect(r.whiteInRim, `${r.name}: white pixels inside the ring`).toBe(0);
      // 2) A continuous beige band wide enough to read as the ring.
      expect(
        r.bestBeigeRun,
        `${r.name}: continuous beige run (>= ${result.minRim} device px)`,
      ).toBeGreaterThanOrEqual(result.minRim);
      // 3) The rim (incl. any lit lip) stays a clear step below page white.
      expect(r.beigeMax, `${r.name}: rim brightness must stay below white`).toBeLessThan(250);
    }
  });
});
