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
