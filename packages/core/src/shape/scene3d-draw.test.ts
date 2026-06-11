import { describe, it, expect } from 'vitest';
import { drawProjected } from './scene3d-draw';
import { computeScene3dQuad, type Vec2 } from './scene3d-camera';

/**
 * Recording 2D context that captures setTransform + drawImage calls so we can
 * assert the mesh-warp behaviour without a real canvas. drawProjected is pure
 * geometry + compositing, so the op sequence is the contract under test.
 */
class RecordingCtx {
  transforms: number[][] = [];
  draws: Array<{ sx: number; sy: number; sw: number; sh: number; dx: number; dy: number; dw: number; dh: number }> = [];
  clips = 0;
  savedDepth = 0;
  maxSavedDepth = 0;

  save(): void {
    this.savedDepth++;
    this.maxSavedDepth = Math.max(this.maxSavedDepth, this.savedDepth);
  }
  restore(): void {
    this.savedDepth--;
  }
  beginPath(): void {}
  moveTo(): void {}
  lineTo(): void {}
  closePath(): void {}
  clip(): void {
    this.clips++;
  }
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.transforms.push([a, b, c, d, e, f]);
  }
  drawImage(
    _img: unknown,
    sx: number,
    sy: number,
    sw: number,
    sh: number,
    dx: number,
    dy: number,
    dw: number,
    dh: number,
  ): void {
    this.draws.push({ sx, sy, sw, sh, dx, dy, dw, dh });
  }
}

function asCtx(c: RecordingCtx): CanvasRenderingContext2D {
  return c as unknown as CanvasRenderingContext2D;
}

const fakeImage = {} as CanvasImageSource;

describe('drawProjected', () => {
  it('draws an identity rectangle as a single un-subdivided affine cell', () => {
    const ctx = new RecordingCtx();
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 0, y: 0 },
      { x: 100, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    drawProjected(fakeImage, asCtx(ctx), 100, 80, corners);
    // Identity quad has zero perspective error → no subdivision → one cell.
    expect(ctx.draws.length).toBe(1);
    expect(ctx.clips).toBe(1);
    // The single cell's transform should be (approximately) the identity scale.
    const [a, b, c, d] = ctx.transforms[0];
    expect(a).toBeCloseTo(1, 3);
    expect(b).toBeCloseTo(0, 3);
    expect(c).toBeCloseTo(0, 3);
    expect(d).toBeCloseTo(1, 3);
  });

  it('subdivides a perspective quad into many cells (error-driven mesh)', () => {
    const ctx = new RecordingCtx();
    // A strongly non-affine quad (foreshortened top edge).
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 30, y: 0 },
      { x: 70, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    drawProjected(fakeImage, asCtx(ctx), 100, 80, corners, 0.5);
    // Must subdivide well past a single cell to hold sub-pixel error.
    expect(ctx.draws.length).toBeGreaterThan(4);
  });

  it('finer tolerance produces at least as many cells as a coarser one', () => {
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 30, y: 0 },
      { x: 70, y: 0 },
      { x: 100, y: 80 },
      { x: 0, y: 80 },
    ];
    const coarse = new RecordingCtx();
    const fine = new RecordingCtx();
    drawProjected(fakeImage, asCtx(coarse), 100, 80, corners, 2);
    drawProjected(fakeImage, asCtx(fine), 100, 80, corners, 0.25);
    expect(fine.draws.length).toBeGreaterThanOrEqual(coarse.draws.length);
  });

  it('skips a degenerate (zero-area) quad without drawing', () => {
    const ctx = new RecordingCtx();
    const corners: [Vec2, Vec2, Vec2, Vec2] = [
      { x: 10, y: 10 },
      { x: 10, y: 10 },
      { x: 10, y: 10 },
      { x: 10, y: 10 },
    ];
    drawProjected(fakeImage, asCtx(ctx), 100, 80, corners);
    expect(ctx.draws.length).toBe(0);
  });

  it('end-to-end: projects a camera quad onto cells covering the image', () => {
    const ctx = new RecordingCtx();
    const q = computeScene3dQuad(
      { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 347 } },
      200,
      150,
    );
    drawProjected(fakeImage, asCtx(ctx), 200, 150, q.corners);
    expect(ctx.draws.length).toBeGreaterThan(1);
    // Every cell samples within the source image bounds (the bleed clamps).
    for (const d of ctx.draws) {
      expect(d.sx).toBeGreaterThanOrEqual(-1e-6);
      expect(d.sy).toBeGreaterThanOrEqual(-1e-6);
      expect(d.sx + d.sw).toBeLessThanOrEqual(200 + 1e-6);
      expect(d.sy + d.sh).toBeLessThanOrEqual(150 + 1e-6);
    }
  });
});
