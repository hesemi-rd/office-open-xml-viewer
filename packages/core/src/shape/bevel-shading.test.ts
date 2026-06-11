import { describe, it, expect } from 'vitest';
import {
  bevelHeightProfile,
  edt1d,
  computeBevelNormals,
  shadePixel,
  lightDirFromRig,
  type BevelShadeParams,
} from './bevel-shading';

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

function normalize(x: number, y: number, z: number) {
  const m = Math.hypot(x, y, z) || 1;
  return { x: x / m, y: y / m, z: z / m };
}
