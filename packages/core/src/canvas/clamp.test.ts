import { describe, it, expect } from 'vitest';
import {
  clampCanvasSize,
  MAX_CANVAS_DIMENSION,
  MAX_CANVAS_AREA,
} from './clamp.js';

// ── clampCanvasSize (RB5): bound a canvas backing size to browser limits ─────

describe('clampCanvasSize — in-budget sizes pass through unchanged', () => {
  it('leaves a normal page size untouched', () => {
    const r = clampCanvasSize(1240, 1754); // ~A4 at 150 DPI
    expect(r).toEqual({ width: 1240, height: 1754, scale: 1, clamped: false });
  });

  it('rounds fractional requests to integers without flagging a clamp', () => {
    const r = clampCanvasSize(800.4, 600.6);
    expect(r.width).toBe(800);
    expect(r.height).toBe(601);
    expect(r.clamped).toBe(false);
    expect(r.scale).toBe(1);
  });

  it('accepts exactly at the per-axis cap', () => {
    const r = clampCanvasSize(MAX_CANVAS_DIMENSION, 1);
    expect(r.width).toBe(MAX_CANVAS_DIMENSION);
    expect(r.height).toBe(1);
    expect(r.clamped).toBe(false);
  });

  it('accepts exactly at the area cap', () => {
    const side = Math.sqrt(MAX_CANVAS_AREA); // 4096
    const r = clampCanvasSize(side, side);
    expect(r.clamped).toBe(false);
    expect(r.width * r.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
  });
});

describe('clampCanvasSize — over-dimension requests scale down, aspect preserved', () => {
  it('clamps a single over-wide axis and scales the other by the same factor', () => {
    // 40000 wide × 500 tall = 20 MP: the width exceeds the axis cap AND the area
    // slightly exceeds the 16 MP cap, so the harder (area) factor wins. Both caps
    // must end up satisfied, and the aspect ratio (80:1) preserved.
    const r = clampCanvasSize(40000, 500);
    expect(r.clamped).toBe(true);
    expect(r.width).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
    expect(r.width * r.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
    const inputAspect = 40000 / 500; // 80
    const outAspect = r.width / r.height;
    // Aspect preserved within a small rounding tolerance.
    expect(Math.abs(outAspect - inputAspect)).toBeLessThan(inputAspect * 0.02);
    // The applied scale is whichever cap binds harder.
    const dimScale = MAX_CANVAS_DIMENSION / 40000;
    const areaScale = Math.sqrt(MAX_CANVAS_AREA / (40000 * 500));
    expect(r.scale).toBeCloseTo(Math.min(dimScale, areaScale), 5);
  });

  it('clamps a purely over-axis request (area within budget) by the dimension factor', () => {
    // 40000 × 100 = 4 MP (within the area cap) but width over the axis cap: only
    // the dimension factor applies.
    const r = clampCanvasSize(40000, 100);
    expect(r.clamped).toBe(true);
    expect(r.width).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
    expect(r.scale).toBeCloseTo(MAX_CANVAS_DIMENSION / 40000, 5);
  });
});

describe('clampCanvasSize — over-area requests scale down by sqrt', () => {
  it('clamps a large square within the area cap, preserving squareness', () => {
    // 10000×10000 = 100 MP, way over the 16 MP area cap but each axis < 32767.
    const r = clampCanvasSize(10000, 10000);
    expect(r.clamped).toBe(true);
    expect(r.width * r.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
    // Square in → square out.
    expect(Math.abs(r.width - r.height)).toBeLessThanOrEqual(1);
    // Linear factor is sqrt(cap / area).
    const expected = Math.sqrt(MAX_CANVAS_AREA / (10000 * 10000));
    expect(r.scale).toBeCloseTo(expected, 5);
  });

  it('applies whichever cap binds harder (area vs dimension)', () => {
    // 32000×20000: within the axis cap on both, but 640 MP ≫ area cap. The area
    // factor must win and bring the product under the area cap.
    const r = clampCanvasSize(32000, 20000);
    expect(r.clamped).toBe(true);
    expect(r.width).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
    expect(r.height).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
    expect(r.width * r.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
    const aspectIn = 32000 / 20000;
    expect(Math.abs(r.width / r.height - aspectIn)).toBeLessThan(0.1);
  });
});

describe('clampCanvasSize — degenerate inputs collapse to a 1×1 canvas', () => {
  it('handles zero, negative, NaN and Infinity', () => {
    for (const [w, h] of [
      [0, 100],
      [-10, 50],
      [NaN, 20],
      [Infinity, 30],
      [10, 0],
    ] as [number, number][]) {
      const r = clampCanvasSize(w, h);
      expect(r.width).toBeGreaterThanOrEqual(1);
      expect(r.height).toBeGreaterThanOrEqual(1);
      expect(Number.isInteger(r.width)).toBe(true);
      expect(Number.isInteger(r.height)).toBe(true);
    }
  });

  it('never returns a sub-1px axis even under extreme aspect clamping', () => {
    // A 1×100000000 strip: the area clamp would drive the 1px axis below 1 — it
    // must be floored to 1, not 0 (a 0-width canvas is unusable).
    const r = clampCanvasSize(1, 100_000_000);
    expect(r.width).toBeGreaterThanOrEqual(1);
    expect(r.height).toBeGreaterThanOrEqual(1);
  });

  it('floors a sub-0.5px positive request to 1×1 (Math.round would send it to 0)', () => {
    // RB5 regression: a positive-but-tiny request (e.g. a 0.4px page at a small
    // DPR) rounds to 0 and previously fell through the UNCLAMPED branch as a 0×0
    // backing store, blanking the viewer. It must become a real 1×1 canvas.
    const r = clampCanvasSize(0.4, 0.4);
    expect(r).toEqual({ width: 1, height: 1, scale: 1, clamped: false });
    // Asymmetric sub-pixel request: each axis floored independently to ≥ 1.
    const r2 = clampCanvasSize(0.2, 3.6);
    expect(r2.width).toBe(1);
    expect(r2.height).toBe(4);
    expect(r2.width).toBeGreaterThanOrEqual(1);
    expect(r2.height).toBeGreaterThanOrEqual(1);
  });
});
