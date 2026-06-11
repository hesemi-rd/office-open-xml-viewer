import { describe, it, expect } from 'vitest';
import { renderPresetShape, hasPreset, buildPresetGeometryPath } from './index';

/**
 * Records every path vertex the engine emits so we can assert on the resulting
 * geometry without a real canvas. Only the path-building methods are needed.
 */
function makeRecorder(): {
  ctx: CanvasRenderingContext2D;
  points: Array<{ x: number; y: number }>;
} {
  const points: Array<{ x: number; y: number }> = [];
  const ctx = {
    beginPath() {},
    closePath() {},
    moveTo(x: number, y: number) {
      points.push({ x, y });
    },
    lineTo(x: number, y: number) {
      points.push({ x, y });
    },
    bezierCurveTo(_x1: number, _y1: number, _x2: number, _y2: number, x: number, y: number) {
      points.push({ x, y });
    },
    quadraticCurveTo(_x1: number, _y1: number, x: number, y: number) {
      points.push({ x, y });
    },
    ellipse(
      cx: number,
      cy: number,
      rx: number,
      ry: number,
      _rot: number,
      _start: number,
      _end: number,
    ) {
      points.push({ x: cx + rx, y: cy });
      points.push({ x: cx - rx, y: cy });
      points.push({ x: cx, y: cy + ry });
      points.push({ x: cx, y: cy - ry });
    },
    save() {},
    restore() {},
    fill() {},
    stroke() {},
    set fillStyle(_v: unknown) {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, points };
}

/** Distinct rounded x/y coordinates touched by the path. */
function distinct(points: Array<{ x: number; y: number }>) {
  const xs = new Set(points.map((p) => Math.round(p.x)));
  const ys = new Set(points.map((p) => Math.round(p.y)));
  return { xs, ys };
}

describe('renderPresetShape (core preset-geometry engine)', () => {
  it('knows the common xlsx preset names', () => {
    expect(hasPreset('parallelogram')).toBe(true);
    expect(hasPreset('rtTriangle')).toBe(true);
    expect(hasPreset('wedgeRectCallout')).toBe(true);
    expect(hasPreset('totallyMadeUpShape')).toBe(false);
  });

  it('renders a parallelogram with slanted (non-rectangular) vertices', () => {
    const { ctx, points } = makeRecorder();
    const ok = renderPresetShape(
      ctx,
      'parallelogram',
      0,
      0,
      200,
      100,
      [], // default adjust
      '#000',
      null,
      () => {},
    );
    expect(ok).toBe(true);
    // A rectangle would touch only x∈{0,200}. A parallelogram's top and bottom
    // edges are offset, so at least one vertex sits strictly between the left
    // and right edges.
    const innerX = points.some((p) => p.x > 1 && p.x < 199);
    expect(innerX).toBe(true);
    // Spans the full box bounds.
    const { xs, ys } = distinct(points);
    expect(Math.max(...xs)).toBeGreaterThanOrEqual(199);
    expect(Math.max(...ys)).toBeGreaterThanOrEqual(99);
  });

  it('renders a right triangle (rtTriangle) as three corners, not a rect', () => {
    const { ctx, points } = makeRecorder();
    const ok = renderPresetShape(
      ctx,
      'rtTriangle',
      0,
      0,
      200,
      100,
      [],
      '#000',
      null,
      () => {},
    );
    expect(ok).toBe(true);
    // rtTriangle: (0,h) (0,0) (w,h). The top-right corner (w,0) must be absent.
    const hasTopRight = points.some((p) => p.x > 199 && p.y < 1);
    expect(hasTopRight).toBe(false);
  });

  it('returns false for an unknown preset so the caller can fall back', () => {
    const { ctx } = makeRecorder();
    const ok = renderPresetShape(
      ctx,
      'totallyMadeUpShape',
      0,
      0,
      10,
      10,
      [],
      '#000',
      null,
      () => {},
    );
    expect(ok).toBe(false);
  });

  // buildPresetGeometryPath — the geometry-only entry used for picture clip
  // silhouettes (ECMA-376 §20.1.9.18). It emits each preset path as a subpath
  // of the current path with no fill/stroke.
  describe('buildPresetGeometryPath (picture clip silhouette)', () => {
    it('builds an ellipse spanning the full box for prst="ellipse"', () => {
      const { ctx, points } = makeRecorder();
      const ok = buildPresetGeometryPath(ctx, 'ellipse', 10, 20, 200, 300);
      expect(ok).toBe(true);
      // The recorder turns ctx.ellipse into its 4 axis-extreme points. The
      // ellipse must be inscribed in the bbox: x∈{10,210}, y∈{20,320}, centred.
      const { xs, ys } = distinct(points);
      expect(Math.min(...xs)).toBe(10);
      expect(Math.max(...xs)).toBe(210);
      expect(Math.min(...ys)).toBe(20);
      expect(Math.max(...ys)).toBe(320);
      // The corners of the bbox must NOT be on the path (an ellipse, not a rect).
      const hasCorner = points.some((p) => Math.round(p.x) === 10 && Math.round(p.y) === 20);
      expect(hasCorner).toBe(false);
    });

    it('builds a rounded-rect silhouette whose corner radius tracks the adjust', () => {
      // roundRect's rounded corners pull the path inward from the box corners by
      // the radius; a larger adjust → vertices further from the corner.
      const inset = (adj: number) => {
        const { ctx, points } = makeRecorder();
        buildPresetGeometryPath(ctx, 'roundRect', 0, 0, 200, 200, [adj]);
        // Smallest non-zero x touched along the top edge ≈ corner radius.
        return Math.min(...points.map((p) => Math.round(p.x)).filter((x) => x > 0));
      };
      expect(inset(40000)).toBeGreaterThan(inset(10000));
    });

    it('returns false for an unknown preset so the caller can fall back to rect', () => {
      const { ctx, points } = makeRecorder();
      const ok = buildPresetGeometryPath(ctx, 'totallyMadeUpShape', 0, 0, 10, 10);
      expect(ok).toBe(false);
      expect(points.length).toBe(0);
    });
  });

  it('honours an explicit adjust value (parallelogram skew)', () => {
    const slant = (adj: number) => {
      const { ctx, points } = makeRecorder();
      renderPresetShape(ctx, 'parallelogram', 0, 0, 200, 100, [adj], '#000', null, () => {});
      // The interior break point along the top/bottom edge moves with adj.
      return points
        .map((p) => Math.round(p.x))
        .filter((x) => x > 1 && x < 199)
        .sort((a, b) => a - b);
    };
    const small = slant(10000);
    const large = slant(50000);
    expect(small).not.toEqual(large);
  });
});
