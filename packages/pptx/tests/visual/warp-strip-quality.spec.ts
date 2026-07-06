/**
 * Browser-rendering quality gate for the WordArt paired-edge warp strips
 * (ECMA-376 §20.1.9.19; piecewise-affine subdivision from #872).
 *
 * The strip draw is exercised in a REAL browser because the artifacts this
 * guards against are browser-specific: Chrome's drawImage/AA pipeline exposed
 * three defects of the original bitmap-blit strips that skia-canvas (the node
 * probe harness) does not reproduce — detached hairlines from stretched
 * source-edge antialiasing, glyphs cut by the offscreen slab rectangle, and
 * blur from re-magnifying a rasterised glyph (vScale up to ~4). The vector
 * clip+fillText strip draw removes the bitmap from the path entirely; these
 * metrics pin that quality in the browser:
 *
 *  - strayInk: faint ink pixels with no solid ink within 3 px — detached
 *    hairlines / stretched AA fringes. Must be 0.
 *  - edgeRatio: (faint edge pixels) / (solid pixels) — a proxy for edge
 *    softness. Warped text must stay in the same regime as flat fillText
 *    (bitmap re-magnification measured ≈2× the flat ratio; vector ≈1×).
 *  - contour steps: P95 of adjacent-column top-contour jumps — staircase
 *    seams at strip boundaries on the sheared Cascade envelope.
 *  - CirclePour ink components: readability guard for high-curvature rings
 *    (radial-sliver regression would shatter glyphs into dozens of parts).
 *
 * Synthetic slides only (no .pptx fetch, no private fixtures) — see
 * warp-quality-fixture.html.
 */
import { test, expect } from '@playwright/test';

// The artifacts under test are HiDPI re-sampling artifacts: render at dpr 2.
test.use({ deviceScaleFactor: 2, viewport: { width: 1400, height: 1600 } });

interface CanvasMetrics {
  solid: number;
  faint: number;
  stray: number;
  components: number;
  contourP95: number;
  contourMax: number;
}

/** Compute ink metrics for one fixture canvas, in-page. */
async function metricsFor(page: import('@playwright/test').Page, id: string): Promise<CanvasMetrics> {
  return page.evaluate((canvasId) => {
    const canvas = document.getElementById(canvasId) as HTMLCanvasElement;
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
    const W = canvas.width;
    const H = canvas.height;
    const data = ctx.getImageData(0, 0, W, H).data;
    // Luminance against the white page background (canvas is opaque white).
    const lum = (i: number) => 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    const solidMask = new Uint8Array(W * H);
    const inkMask = new Uint8Array(W * H); // any visible ink, however faint
    let solid = 0;
    let faint = 0;
    for (let p = 0; p < W * H; p++) {
      const l = lum(p * 4);
      if (l <= 100) { solidMask[p] = 1; inkMask[p] = 1; solid++; }
      else if (l < 235) { inkMask[p] = 1; if (l > 100 && l < 220) faint++; }
    }
    // Stray faint ink: visible ink with no solid ink within Chebyshev radius 3.
    let stray = 0;
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const p = y * W + x;
        if (!inkMask[p] || solidMask[p]) continue;
        let nearSolid = false;
        for (let dy = -3; dy <= 3 && !nearSolid; dy++) {
          const yy = y + dy;
          if (yy < 0 || yy >= H) continue;
          for (let dx = -3; dx <= 3; dx++) {
            const xx = x + dx;
            if (xx < 0 || xx >= W) continue;
            if (solidMask[yy * W + xx]) { nearSolid = true; break; }
          }
        }
        if (!nearSolid) stray++;
      }
    }
    // 8-connected components of VISIBLE ink (readability / sliver guard; the
    // faint tier is included so a hairline-thin but continuous stroke — e.g.
    // the compressed inner radius of a ring — counts as one piece, matching
    // the node/skia probe's ink threshold).
    const seen = new Uint8Array(W * H);
    const stack: number[] = [];
    let components = 0;
    for (let start = 0; start < W * H; start++) {
      if (!inkMask[start] || seen[start]) continue;
      components++;
      stack.length = 0;
      stack.push(start);
      seen[start] = 1;
      while (stack.length) {
        const p = stack.pop() as number;
        const x = p % W, y = (p / W) | 0;
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          if (!dx && !dy) continue;
          const nx = x + dx, ny = y + dy;
          if (nx < 0 || ny < 0 || nx >= W || ny >= H) continue;
          const np = ny * W + nx;
          if (inkMask[np] && !seen[np]) { seen[np] = 1; stack.push(np); }
        }
      }
    }
    // Top-contour steps: per column, topmost row with lum<160; jumps between
    // adjacent inked columns. P95 filters the rare legitimate intra-glyph
    // feature jump; strip-boundary staircases inflate the bulk of the
    // distribution instead.
    const tops: Array<number | null> = new Array(W).fill(null);
    for (let x = 0; x < W; x++) {
      for (let y = 0; y < H; y++) {
        if (lum((y * W + x) * 4) < 160) { tops[x] = y; break; }
      }
    }
    const jumps: number[] = [];
    for (let x = 1; x < W; x++) {
      const a = tops[x - 1];
      const b = tops[x];
      if (a != null && b != null) jumps.push(Math.abs(b - a));
    }
    jumps.sort((m, n) => m - n);
    const contourP95 = jumps.length ? jumps[Math.floor(jumps.length * 0.95)] : 0;
    const contourMax = jumps.length ? jumps[jumps.length - 1] : 0;
    return { solid, faint, stray, components, contourP95, contourMax };
  }, id);
}

test('warp strips render clean in a real browser (stray ink / sharpness / seams)', async ({ page }) => {
  await page.goto('/tests/visual/warp-quality-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready', undefined, {
    timeout: 30_000,
  });

  const flat = await metricsFor(page, 'flat');
  const results: Record<string, CanvasMetrics> = { flat };
  for (const id of ['inflate', 'deflate', 'cascade', 'wave1', 'circlepour']) {
    results[id] = await metricsFor(page, id);
  }
  // eslint-disable-next-line no-console
  console.log('[warp-quality]', JSON.stringify(results, null, 1));

  const flatRatio = flat.faint / Math.max(1, flat.solid);
  for (const id of ['inflate', 'deflate', 'cascade', 'wave1', 'circlepour']) {
    const m = results[id];
    expect(m.solid, `${id}: has ink`).toBeGreaterThan(500);
    // 1) No detached hairlines / stretched AA fringes anywhere.
    expect(m.stray, `${id}: stray faint ink pixels`).toBe(0);
    // 2) Edge softness stays in the flat-fillText regime. The warp stretches
    //    glyphs (vScale up to ~4 → up to ~4× the outline length per solid px),
    //    so allow that geometric factor but not the compounding blur of a
    //    re-magnified bitmap (measured ≈2× worse again on the bitmap path).
    expect(m.faint / Math.max(1, m.solid), `${id}: edge-softness ratio`).toBeLessThan(
      flatRatio * 4,
    );
  }
  // 3) Sheared piecewise-linear envelope must not staircase at strip seams:
  //    the bulk of adjacent-column contour jumps stays at the AA scale.
  expect(results.cascade.contourP95, 'cascade: top-contour P95 step').toBeLessThanOrEqual(3);
  // 4) High-curvature ring stays readable — one blob per glyph part, not a
  //    fan of radial slivers ("Round" = 5 single-piece letters; warped letters
  //    may merge or split at hairline joins, slivers are dozens).
  expect(results.circlepour.components, 'circlepour: ink components').toBeLessThanOrEqual(10);
});
