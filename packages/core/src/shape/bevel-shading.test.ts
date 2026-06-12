import { describe, it, expect } from 'vitest';
import {
  bevelHeightProfile,
  edt1d,
  computeBevelNormals,
  shadePixel,
  lightDirFromRig,
  applyBevelShading,
  applyExtrusion,
  type BevelShadeParams,
  type BevelCtx,
} from './bevel-shading';

/** Minimal ImageData-backed BevelCtx for the compositor functions. */
function fakeCtx(w: number, h: number, fill?: (x: number, y: number) => [number, number, number, number]): BevelCtx & { data: Uint8ClampedArray } {
  const data = new Uint8ClampedArray(w * h * 4);
  if (fill) {
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const [r, g, b, a] = fill(x, y);
        const o = (y * w + x) * 4;
        data[o] = r; data[o + 1] = g; data[o + 2] = b; data[o + 3] = a;
      }
  }
  return {
    data,
    canvas: { width: w, height: h },
    getImageData() {
      return { data: data.slice(), width: w, height: h } as unknown as ImageData;
    },
    putImageData(img: ImageData) {
      data.set(img.data);
    },
  };
}

describe('bevelHeightProfile', () => {
  it('circle profile rises as a quarter-circle: h(0)=0, h(w)=1, convex', () => {
    const w = 10;
    const p = bevelHeightProfile('circle', w);
    expect(p(0)).toBeCloseTo(0, 5);
    expect(p(w)).toBeCloseTo(1, 5);
    // Quarter circle h = sqrt(1-(1-t)^2): at t=0.5 → sqrt(1-0.25)=0.866 (convex,
    // above the linear 0.5).
    expect(p(w * 0.5)).toBeCloseTo(Math.sqrt(1 - 0.25), 5);
    expect(p(w * 0.5)).toBeGreaterThan(0.5);
  });

  it('hardEdge is a linear chamfer: h(t) = t/w', () => {
    const w = 8;
    const p = bevelHeightProfile('hardEdge', w);
    expect(p(0)).toBeCloseTo(0, 5);
    expect(p(w)).toBeCloseTo(1, 5);
    expect(p(w * 0.5)).toBeCloseTo(0.5, 5);
  });

  it('relaxedInset (default) is a smooth S-curve in [0,1]', () => {
    const w = 10;
    const p = bevelHeightProfile('relaxedInset', w);
    expect(p(0)).toBeCloseTo(0, 5);
    expect(p(w)).toBeCloseTo(1, 5);
    // Monotonic increasing.
    expect(p(w * 0.25)).toBeLessThan(p(w * 0.75));
  });

  it('clamps distance beyond the band to the flat top (h=1)', () => {
    const w = 5;
    const p = bevelHeightProfile('circle', w);
    expect(p(w * 2)).toBeCloseTo(1, 5);
  });

  it('zero width yields a flat profile (always 1, no bevel)', () => {
    const p = bevelHeightProfile('circle', 0);
    expect(p(0)).toBeCloseTo(1, 5);
    expect(p(10)).toBeCloseTo(1, 5);
  });
});

describe('edt1d (squared Euclidean distance transform, Felzenszwalb)', () => {
  it('computes squared distance to the nearest zero on a 1D line', () => {
    // f: 0 at index 2, +inf elsewhere → squared distance to index 2.
    const INF = 1e20;
    const f = [INF, INF, 0, INF, INF];
    const out = edt1d(f);
    expect(Array.from(out)).toEqual([4, 1, 0, 1, 4]);
  });

  it('handles two seeds, taking the nearer', () => {
    const INF = 1e20;
    const f = [0, INF, INF, INF, 0];
    const out = edt1d(f);
    expect(Array.from(out)).toEqual([0, 1, 4, 1, 0]);
  });
});

describe('lightDirFromRig', () => {
  it('dir="t" points the light from the top (negative screen-y), toward viewer', () => {
    const L = lightDirFromRig('threePt', 't');
    expect(L.y).toBeLessThan(0); // light comes from above (screen y down)
    expect(L.z).toBeGreaterThan(0); // and from in front of the surface
    // unit length
    expect(Math.hypot(L.x, L.y, L.z)).toBeCloseTo(1, 5);
  });

  it('dir="l" comes from the left (negative x)', () => {
    const L = lightDirFromRig('threePt', 'l');
    expect(L.x).toBeLessThan(0);
  });

  it('dir="br" comes from bottom-right (positive x, positive y)', () => {
    const L = lightDirFromRig('threePt', 'br');
    expect(L.x).toBeGreaterThan(0);
    expect(L.y).toBeGreaterThan(0);
  });
});

describe('shadePixel (Lambert + weak specular)', () => {
  const params: BevelShadeParams = {
    light: lightDirFromRig('threePt', 't'),
    material: 'matte',
    ambient: 0.55,
    diffuse: 0.45,
    specular: 0.0,
    shininess: 8,
  };

  it('a flat top-facing normal (0,0,1) gets a mid factor near ambient+diffuse·(L·N)', () => {
    const f = shadePixel({ x: 0, y: 0, z: 1 }, params);
    // L·N = L.z. factor = ambient + diffuse*max(0,L.z).
    const expected = params.ambient + params.diffuse * params.light.z;
    expect(f).toBeCloseTo(expected, 5);
  });

  it('a normal tilted toward the light is brighter than one tilted away', () => {
    const toward = shadePixel(normalize(params.light.x, params.light.y, 1), params);
    const away = shadePixel(normalize(-params.light.x, -params.light.y, 1), params);
    expect(toward).toBeGreaterThan(away);
  });

  it('plastic material adds a specular highlight (brighter peak than matte)', () => {
    const matteP = { ...params, specular: 0.0 };
    const plasticP = { ...params, specular: 0.4, material: 'plastic' as const };
    // Normal exactly on the light direction → max specular.
    const n = normalize(params.light.x, params.light.y, params.light.z);
    expect(shadePixel(n, plasticP)).toBeGreaterThan(shadePixel(n, matteP));
  });

  it('clamps the factor to >= 0 (no negative)', () => {
    const f = shadePixel({ x: 0, y: 0, z: -1 }, params);
    expect(f).toBeGreaterThanOrEqual(0);
  });
});

describe('computeBevelNormals (finite-difference of height field)', () => {
  it('produces (0,0,1) on the flat interior and tilted normals on the band', () => {
    // 7x7 filled square, bevel width 2. Interior pixels (far from edge) flat;
    // band pixels tilt outward.
    const W = 7;
    const H = 7;
    const alpha = new Uint8ClampedArray(W * H).fill(255);
    const { normals, bandMask } = computeBevelNormals(alpha, W, H, 2, 'circle', 1);
    // Centre pixel (3,3) is interior → flat normal.
    const ci = (3 * W + 3) * 3;
    expect(normals[ci + 2]).toBeGreaterThan(0.9); // nz ≈ 1
    expect(bandMask[3 * W + 3]).toBe(0); // not in band
    // An edge-band pixel (e.g. (3,0), on the top edge, distance 1 < band 2) is
    // in the band and its normal tilts upward (ny < 0 → faces up toward the top).
    const ei = (0 * W + 3) * 3;
    expect(bandMask[0 * W + 3]).toBe(1);
    expect(normals[ei + 1]).toBeLessThan(0); // ny<0 faces up/out at the top edge
  });

  it('ignores fully transparent pixels (outside the silhouette)', () => {
    const W = 5;
    const H = 5;
    const alpha = new Uint8ClampedArray(W * H); // all zero
    const { bandMask } = computeBevelNormals(alpha, W, H, 2, 'circle', 1);
    expect(bandMask.every((v) => v === 0)).toBe(true);
  });
});

/**
 * Anti-aliased ellipse coverage mask (mimics a Canvas `clip()` rasterisation:
 * interior 255, exterior 0, a 1-px fractional rim). Used to exercise the bevel
 * lip on a curved silhouette with a wide band — the case that exposed the
 * EDT-gradient faceting regression (PR #410 → this fix).
 */
function ellipseAlpha(W: number, H: number, ss = 4): Uint8ClampedArray {
  const a = new Uint8ClampedArray(W * H);
  const cx = W / 2;
  const cy = H / 2;
  const rx = W / 2 - 1;
  const ry = H / 2 - 1;
  for (let y = 0; y < H; y++)
    for (let x = 0; x < W; x++) {
      let cnt = 0;
      for (let sy = 0; sy < ss; sy++)
        for (let sx = 0; sx < ss; sx++) {
          const px = x + (sx + 0.5) / ss - 0.5;
          const py = y + (sy + 0.5) / ss - 0.5;
          if (((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1) cnt++;
        }
      a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
    }
  return a;
}

describe('computeBevelNormals — smooth lip on a curved silhouette (anti-facet)', () => {
  // REGRESSION (PR #410): a wide hardEdge bevel on an ellipse produced a faceted
  // (polygonal) lip because the per-pixel finite-difference gradient of the EDT
  // height field is piecewise-constant — each band pixel's distance is dominated
  // by ONE nearest boundary sample, so ∇d snaps to the Voronoi-cell direction of
  // the silhouette's sampled rim. The lip normal must instead track the
  // *macroscopic* silhouette direction so a circle reads as a circle, not an
  // N-gon. We assert that the in-plane normal azimuth varies MONOTONICALLY and
  // SMOOTHLY around the ring, with no large angular jumps between neighbouring
  // boundary samples (a facet shows up as a long run of identical azimuths broken
  // by a sudden kink).
  it('in-plane normal azimuth rotates smoothly around a wide-band ellipse', () => {
    const W = 240;
    const H = 320;
    const band = 40; // wide band — the regime where faceting was visible
    const alpha = ellipseAlpha(W, H);
    const { normals, bandMask } = computeBevelNormals(alpha, W, H, band, 'hardEdge', band / 2);

    // Walk a ring near the OUTER rim (just inside the silhouette) sampling the
    // in-plane normal azimuth at many angles; it should rotate ~once around 2π.
    const cx = W / 2;
    const cy = H / 2;
    const rx = W / 2 - 1;
    const ry = H / 2 - 1;
    // Sample along the band MIDLINE (~band/2 inside the rim) where the lip faces
    // the camera at its steepest and where the facets were most visible. The very
    // outermost 1-2 px are the anti-aliased rim fringe (sub-pixel coverage noise),
    // which is not a facet, so we measure inside it.
    const inset = band / 2;
    const azimuths: number[] = [];
    for (let deg = 0; deg < 360; deg += 1) {
      const ang = (deg * Math.PI) / 180;
      const ex = Math.cos(ang) * (rx - inset);
      const ey = Math.sin(ang) * (ry - inset);
      const x = Math.round(cx + ex);
      const y = Math.round(cy + ey);
      if (x < 0 || y < 0 || x >= W || y >= H) continue;
      const i = y * W + x;
      if (!bandMask[i]) continue;
      const nx = normals[i * 3];
      const ny = normals[i * 3 + 1];
      if (Math.hypot(nx, ny) < 1e-3) continue;
      azimuths.push(Math.atan2(ny, nx));
    }
    expect(azimuths.length).toBeGreaterThan(300);

    // Unwrap and measure the largest jump between consecutive samples. For a
    // smooth lip the azimuth advances by ~(2π / N) ≈ 0.017 rad per degree-step;
    // a facet would hold one azimuth for many steps then jump by a big angle.
    // Allow up to ~12° (0.21 rad) between adjacent 1° samples — generous enough
    // for the rim discretisation yet far below the >30° jumps the facets caused.
    let maxJump = 0;
    for (let i = 1; i < azimuths.length; i++) {
      let d = azimuths[i] - azimuths[i - 1];
      while (d > Math.PI) d -= 2 * Math.PI;
      while (d < -Math.PI) d += 2 * Math.PI;
      maxJump = Math.max(maxJump, Math.abs(d));
    }
    expect(maxJump).toBeLessThan(0.21);

    // And the azimuth must make ~one full turn (net rotation ≈ 2π), i.e. the lip
    // faces radially outward all the way round — not collapse onto a few facet
    // directions.
    let net = 0;
    for (let i = 1; i < azimuths.length; i++) {
      let d = azimuths[i] - azimuths[i - 1];
      while (d > Math.PI) d -= 2 * Math.PI;
      while (d < -Math.PI) d += 2 * Math.PI;
      net += d;
    }
    expect(Math.abs(Math.abs(net) - 2 * Math.PI)).toBeLessThan(0.5);
  });
});

describe('applyExtrusion (side-wall sweep, §20.1.5.12)', () => {
  it('fills the swept band behind the front face with the wall colour', () => {
    // 20x20 with an opaque 8x8 block at (6,6)-(13,13). Push it right+down by
    // (4,4): the side wall is the L-shaped band uncovered to the right/below.
    const W = 20;
    const H = 20;
    const ctx = fakeCtx(W, H, (x, y) =>
      x >= 6 && x < 14 && y >= 6 && y < 14 ? [10, 20, 200, 255] : [0, 0, 0, 0],
    );
    applyExtrusion(ctx, { offsetX: 4, offsetY: 4, rgb: [50, 50, 50] });
    // A pixel just below-right of the block (e.g. (15,15)) should now be wall.
    const wall = (15 * W + 15) * 4;
    expect(ctx.data[wall + 3]).toBe(255);
    expect(ctx.data[wall]).toBe(50);
    // The front face is untouched.
    const face = (8 * W + 8) * 4;
    expect(ctx.data[face]).toBe(10);
    expect(ctx.data[face + 2]).toBe(200);
  });

  it('is a no-op for a sub-pixel offset (face-on camera)', () => {
    const W = 10;
    const H = 10;
    const ctx = fakeCtx(W, H, (x, y) => (x >= 3 && x < 7 && y >= 3 && y < 7 ? [10, 10, 10, 255] : [0, 0, 0, 0]));
    const before = ctx.data.slice();
    applyExtrusion(ctx, { offsetX: 0.2, offsetY: 0.1, rgb: [99, 99, 99] });
    expect(Array.from(ctx.data)).toEqual(Array.from(before));
  });
});

describe('applyBevelShading (end-to-end on a filled square)', () => {
  it('brightens the top lip and darkens the bottom lip under a top light', () => {
    const W = 24;
    const H = 24;
    // Mid-grey opaque square filling the canvas.
    const ctx = fakeCtx(W, H, () => [120, 120, 120, 255]);
    applyBevelShading(ctx, {
      widthPx: 5,
      heightPx: 4,
      prst: 'circle',
      material: 'matte',
      light: lightDirFromRig('threePt', 't'),
    });
    const lumAt = (x: number, y: number) => {
      const o = (y * W + x) * 4;
      return ctx.data[o] * 0.299 + ctx.data[o + 1] * 0.587 + ctx.data[o + 2] * 0.114;
    };
    // Top lip (y=1) lit brighter than the flat centre; bottom lip (y=22) darker.
    expect(lumAt(12, 1)).toBeGreaterThan(lumAt(12, 12));
    expect(lumAt(12, 22)).toBeLessThan(lumAt(12, 12));
  });
});

function normalize(x: number, y: number, z: number) {
  const m = Math.hypot(x, y, z) || 1;
  return { x: x / m, y: y / m, z: z / m };
}
