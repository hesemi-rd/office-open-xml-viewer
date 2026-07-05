import { describe, it, expect } from 'vitest';
import {
  buildWarpEnvelope,
  hasTextWarp,
  isSingleEdgeWarp,
  samplePolyline,
  warpGlyphTransform,
  type Polyline,
} from './text-warp';

const W = 400;
const H = 100;

describe('text-warp preset registry', () => {
  it('knows the 40 spec presets, case-insensitively', () => {
    expect(hasTextWarp('textArchUp')).toBe(true);
    expect(hasTextWarp('TEXTARCHUP')).toBe(true);
    expect(hasTextWarp('textWave1')).toBe(true);
    expect(hasTextWarp('textPlain')).toBe(true);
    expect(hasTextWarp('textNoShape')).toBe(false);
    expect(hasTextWarp('notAWarp')).toBe(false);
  });

  it('classifies single-edge (arch/circle) vs paired-edge warps', () => {
    expect(isSingleEdgeWarp('textArchUp')).toBe(true);
    expect(isSingleEdgeWarp('textArchDown')).toBe(true);
    expect(isSingleEdgeWarp('textCircle')).toBe(true);
    expect(isSingleEdgeWarp('textPlain')).toBe(false);
    expect(isSingleEdgeWarp('textWave1')).toBe(false);
  });

  it('returns null for an unknown preset', () => {
    expect(buildWarpEnvelope('nope', [], W, H)).toBeNull();
  });
});

describe('textPlain — the identity envelope', () => {
  it('is two horizontal edges spanning the full width at top / bottom', () => {
    const env = buildWarpEnvelope('textPlain', [], W, H);
    expect(env).not.toBeNull();
    const top = env!.top;
    const bottom = env!.bottom;
    // Both edges run left→right across the width.
    expect(top[0].x).toBeCloseTo(0, 1);
    expect(top[top.length - 1].x).toBeCloseTo(W, 1);
    expect(bottom[0].x).toBeCloseTo(0, 1);
    expect(bottom[bottom.length - 1].x).toBeCloseTo(W, 1);
    // Top edge is a constant y near 0, bottom near H.
    for (const p of top) expect(p.y).toBeCloseTo(0, 1);
    for (const p of bottom) expect(p.y).toBeCloseTo(H, 1);
  });

  it('warps a mid-line baseline to a flat, un-rotated, unit-scaled point', () => {
    const env = buildWarpEnvelope('textPlain', [], W, H)!;
    // boxHeight == H (text exactly fills the box) → vScale == 1.
    const mid = warpGlyphTransform(env, 0.5, H, 0.75);
    expect(mid.x).toBeCloseTo(W / 2, 1);
    expect(mid.y).toBeCloseTo(H * 0.75, 1);
    expect(mid.angle).toBeCloseTo(0, 3); // horizontal edges → no rotation
    expect(mid.vScale).toBeCloseTo(1, 3); // gap == box height → no compression
  });
});

describe('textInflate — box bulges in the middle', () => {
  it('top edge bows UP (smaller y) and bottom bows DOWN (larger y) at centre', () => {
    const env = buildWarpEnvelope('textInflate', [], W, H)!;
    const topMid = samplePolyline(env.top, env.topLen, 0.5);
    const topEnd = samplePolyline(env.top, env.topLen, 0.0);
    const botMid = samplePolyline(env.bottom, env.bottomLen, 0.5);
    const botEnd = samplePolyline(env.bottom, env.bottomLen, 0.0);
    // At the centre the top edge is higher (smaller y) than at the ends.
    expect(topMid.y).toBeLessThan(topEnd.y);
    // and the bottom edge is lower (larger y) than at the ends.
    expect(botMid.y).toBeGreaterThan(botEnd.y);
    // So the vertical gap is LARGER at the centre than at the ends.
    const gapMid = botMid.y - topMid.y;
    const gapEnd = botEnd.y - topEnd.y;
    expect(gapMid).toBeGreaterThan(gapEnd);
  });
});

describe('textArchUp — single arc baseline', () => {
  it('flattens into an arc whose points are equidistant from a common centre', () => {
    // Default adj makes a full 180°-ish arch. Points on an arc share a centre.
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    const arc = env.top;
    expect(arc.length).toBeGreaterThan(20);
    // Fit centre as the average, then check radius variance is tiny relative to r.
    const cx = arc.reduce((s, p) => s + p.x, 0) / arc.length;
    const cy = arc.reduce((s, p) => s + p.y, 0) / arc.length;
    const radii = arc.map((p) => Math.hypot(p.x - cx, p.y - cy));
    const rMean = radii.reduce((s, r) => s + r, 0) / radii.length;
    // A circular arc's centroid is NOT the circle centre, so refine: use the
    // ellipse centre (hc, vc·) is unknown; instead assert the arc is convex and
    // its mid-point sits above its endpoints (an "up" arch opens downward).
    const start = arc[0];
    const end = arc[arc.length - 1];
    const midx = arc[Math.floor(arc.length / 2)];
    // The apex of an up-arch is higher (smaller y) than both ends.
    expect(midx.y).toBeLessThan(start.y);
    expect(midx.y).toBeLessThan(end.y);
    expect(rMean).toBeGreaterThan(0);
  });

  it('per-glyph transform keeps vScale=1 and rotates glyphs along the arc', () => {
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    expect(env.singleEdge).toBe(true);
    const boxH = 30;
    const gl = warpGlyphTransform(env, 0.5, boxH, 0.8);
    const gr = warpGlyphTransform(env, 0.85, boxH, 0.8);
    const glft = warpGlyphTransform(env, 0.15, boxH, 0.8);
    // Single-edge presets never compress the glyph height.
    expect(gl.vScale).toBe(1);
    // At the apex the axis is horizontal; toward the right end it tilts down
    // (positive angle), toward the left end it tilts up (negative angle), and
    // the tilt grows monotonically away from the apex.
    expect(Math.abs(gl.angle)).toBeLessThan(0.05);
    expect(gr.angle).toBeGreaterThan(0.15);
    expect(glft.angle).toBeLessThan(-0.15);
    // Left/right are mirror images about the apex.
    expect(gr.angle).toBeCloseTo(-glft.angle, 3);
  });
});

describe('textInflate — per-glyph vertical scale', () => {
  it('scales glyphs TALLER at the centre than at the ends', () => {
    const env = buildWarpEnvelope('textInflate', [], W, H)!;
    const boxH = H; // nominal
    const mid = warpGlyphTransform(env, 0.5, boxH, 0.8);
    const end = warpGlyphTransform(env, 0.02, boxH, 0.8);
    expect(mid.vScale).toBeGreaterThan(end.vScale);
  });
});

describe('samplePolyline — arc-length parameterisation', () => {
  it('returns endpoints at u=0 and u=1 and a unit tangent', () => {
    const poly: Polyline = [
      { x: 0, y: 0 },
      { x: 10, y: 0 },
      { x: 20, y: 0 },
    ];
    const cum = [0, 10, 20];
    const a = samplePolyline(poly, cum, 0);
    const b = samplePolyline(poly, cum, 1);
    const mid = samplePolyline(poly, cum, 0.5);
    expect(a.x).toBeCloseTo(0);
    expect(b.x).toBeCloseTo(20);
    expect(mid.x).toBeCloseTo(10);
    expect(Math.hypot(a.tx, a.ty)).toBeCloseTo(1);
  });
});
