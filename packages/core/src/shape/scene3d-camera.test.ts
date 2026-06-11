import { describe, it, expect } from 'vitest';
import {
  computeScene3dQuad,
  isScene3dNonIdentity,
  computeDepthOffset,
  type CameraInput,
  type Vec2,
} from './scene3d-camera';

const W = 400;
const H = 300;

/** Round a Vec2 for readable snapshots / approximate comparisons. */
function r(p: Vec2, d = 3): { x: number; y: number } {
  const k = 10 ** d;
  return { x: Math.round(p.x * k) / k, y: Math.round(p.y * k) / k };
}

describe('computeScene3dQuad', () => {
  it('orthographicFront with no rot is the identity rectangle', () => {
    const cam: CameraInput = { prst: 'orthographicFront' };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isIdentity).toBe(true);
    expect(q.isAffine).toBe(true);
    expect(r(q.corners[0])).toEqual({ x: 0, y: 0 });
    expect(r(q.corners[1])).toEqual({ x: W, y: 0 });
    expect(r(q.corners[2])).toEqual({ x: W, y: H });
    expect(r(q.corners[3])).toEqual({ x: 0, y: H });
  });

  it('orthographicFront with all-zero rot is still identity', () => {
    const cam: CameraInput = { prst: 'orthographicFront', rot: { lat: 0, lon: 0, rev: 0 } };
    expect(computeScene3dQuad(cam, W, H).isIdentity).toBe(true);
    expect(isScene3dNonIdentity(cam)).toBe(false);
  });

  it('unknown preset falls back to identity (no throw)', () => {
    const cam: CameraInput = { prst: 'legacyObliqueTopLeft' };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isIdentity).toBe(true);
  });

  it('lon>0 turns the right edge TOWARD the viewer (right edge nearer/taller)', () => {
    // A pure longitude (Y-axis) turn. With lon > 0 the right edge comes toward
    // the viewer and the left edge recedes (file-angle convention — see the
    // module header). The left and right edges stay vertical (equal x for the
    // two corners on each side) and the quad mirrors top↔bottom about the
    // horizontal centre line; the near (right) edge is TALLER than the far (left)
    // edge by perspective foreshortening.
    const cam: CameraInput = { prst: 'perspectiveFront', rot: { lat: 0, lon: 25, rev: 0 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(false);
    const [tl, tr, br, bl] = q.corners.map((c) => r(c));
    // Left edge vertical (TL.x == BL.x); right edge vertical (TR.x == BR.x).
    expect(tl.x).toBeCloseTo(bl.x, 2);
    expect(tr.x).toBeCloseTo(br.x, 2);
    // Mirror about the horizontal centre line: TL.y↔BL.y, TR.y↔BR.y.
    expect(tl.y).toBeCloseTo(H - bl.y, 2);
    expect(tr.y).toBeCloseTo(H - br.y, 2);
    // Right (near) edge taller than left (far) edge.
    const leftH = bl.y - tl.y;
    const rightH = br.y - tr.y;
    expect(rightH).toBeGreaterThan(leftH + 1);
  });

  it('lat>0 tips the TOP edge TOWARD the viewer (top edge nearer/wider)', () => {
    // A pure latitude (X-axis) tilt. With lat > 0 the top edge tips toward the
    // viewer and the bottom recedes (file-angle convention — see the module
    // header; sample-11's lat = −30° does the opposite, top recedes). The top and
    // bottom edges stay horizontal and the quad mirrors left↔right about the
    // vertical centre line; the near (top) edge is WIDER than the far (bottom).
    const cam: CameraInput = { prst: 'perspectiveFront', rot: { lat: 25, lon: 0, rev: 0 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(false);
    const [tl, tr, br, bl] = q.corners.map((c) => r(c));
    // Top edge horizontal (TL.y == TR.y); bottom edge horizontal (BL.y == BR.y).
    expect(tl.y).toBeCloseTo(tr.y, 2);
    expect(bl.y).toBeCloseTo(br.y, 2);
    // Mirror about the vertical centre line: TL.x↔TR.x, BL.x↔BR.x.
    expect(tl.x).toBeCloseTo(W - tr.x, 2);
    expect(bl.x).toBeCloseTo(W - br.x, 2);
    // Top edge WIDER than bottom edge (top nears the viewer).
    const topW = tr.x - tl.x;
    const botW = br.x - bl.x;
    expect(topW).toBeGreaterThan(botW);
  });

  it('lat<0 (sample-11 slide-3 tilt) makes the TOP edge recede (narrower)', () => {
    // sample-11 slide 3 supplies lat = 330° (= −30°). The corrected convention
    // must make the top edge RECEDE here, matching the PDF ground truth
    // (top/bottom width ratio ≈ 0.82). This is the inverse of the lat>0 case and
    // is the regression guard for the axis-convention fix.
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 0, rev: 0 } };
    const q = computeScene3dQuad(cam, W, H);
    const [tl, tr, br, bl] = q.corners;
    const topW = tr.x - tl.x;
    const botW = br.x - bl.x;
    expect(topW).toBeLessThan(botW);
  });

  it('rev-only rotation is a pure in-plane rotation (affine, rigid)', () => {
    // rev spins the shape in the screen plane: the quad stays a rectangle
    // (affine), corner-to-corner distances are preserved up to the refit scale.
    // rev > 0 rotates COUNTER-CLOCKWISE on screen (file-angle convention).
    const cam: CameraInput = { prst: 'orthographicFront', rot: { lat: 0, lon: 0, rev: 30 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(true);
    expect(q.isIdentity).toBe(false);
    // Edge lengths: opposite edges equal, adjacent edges in the original w:h
    // ratio (rigid rotation preserves the rectangle's proportions).
    const [tl, tr, br, bl] = q.corners;
    const top = Math.hypot(tr.x - tl.x, tr.y - tl.y);
    const right = Math.hypot(br.x - tr.x, br.y - tr.y);
    const bottom = Math.hypot(br.x - bl.x, br.y - bl.y);
    const left = Math.hypot(bl.x - tl.x, bl.y - tl.y);
    expect(top).toBeCloseTo(bottom, 1);
    expect(left).toBeCloseTo(right, 1);
    expect(top / right).toBeCloseTo(W / H, 2);
    // Quad centroid stays at the bbox centre.
    const cx = (tl.x + tr.x + br.x + bl.x) / 4;
    const cy = (tl.y + tr.y + br.y + bl.y) / 4;
    expect(cx).toBeCloseTo(W / 2, 3);
    expect(cy).toBeCloseTo(H / 2, 3);
    // Direction: rev > 0 is counter-clockwise on screen, so the top-right corner
    // rises above the top-left corner (TR.y < TL.y in y-down coords).
    expect(tr.y).toBeLessThan(tl.y);
  });

  it('fov override widens the perspective foreshortening', () => {
    const base: CameraInput = { prst: 'perspectiveFront', rot: { lat: 30, lon: 0, rev: 0 } };
    const wide: CameraInput = { ...base, fov: 90 };
    const qb = computeScene3dQuad(base, W, H);
    const qw = computeScene3dQuad(wide, W, H);
    // Wider FOV → stronger near/far size difference → bigger top-vs-bottom width gap.
    const gap = (q: ReturnType<typeof computeScene3dQuad>): number => {
      const topW = q.corners[1].x - q.corners[0].x;
      const botW = q.corners[2].x - q.corners[3].x;
      return Math.abs(topW - botW);
    };
    expect(gap(qw)).toBeGreaterThan(gap(qb));
  });

  it('zoom keeps the quad centred (refit absorbs uniform scale)', () => {
    // The refit pins the projected quad's bounding-box centre to the element's
    // bbox centre, so a uniform zoom does not shift the quad.
    const cam: CameraInput = { prst: 'perspectiveFront', rot: { lat: 20, lon: 10, rev: 0 }, zoom: 2 };
    const q = computeScene3dQuad(cam, W, H);
    const xs = q.corners.map((c) => c.x);
    const ys = q.corners.map((c) => c.y);
    const bbCx = (Math.min(...xs) + Math.max(...xs)) / 2;
    const bbCy = (Math.min(...ys) + Math.max(...ys)) / 2;
    expect(bbCx).toBeCloseTo(W / 2, 3);
    expect(bbCy).toBeCloseTo(H / 2, 3);
  });

  it('projected quad fits inside the original bounding box', () => {
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 347 } };
    const q = computeScene3dQuad(cam, W, H);
    for (const c of q.corners) {
      expect(c.x).toBeGreaterThanOrEqual(-0.001);
      expect(c.x).toBeLessThanOrEqual(W + 0.001);
      expect(c.y).toBeGreaterThanOrEqual(-0.001);
      expect(c.y).toBeLessThanOrEqual(H + 0.001);
    }
  });

  it('snapshot — sample-11 slide-3 camera (perspectiveRelaxed lat330 lon20 rev347)', () => {
    // The exact camera from sample-11 slide 3, "図 3". lat=330° (=−30° tilt → top
    // recedes), lon=20° (right edge nears), rev=347° (=−13° → clockwise in-plane),
    // fov=26° (calibrated, see scene3d-camera.ts). Fixed numeric snapshot so
    // future changes to the camera math are caught; the corner positions encode
    // the corrected axis convention validated against sample-11.pdf page 3.
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 347 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(false);
    // Top edge narrower than bottom edge (top recedes — the PDF keystone).
    const topW = q.corners[1].x - q.corners[0].x;
    const botW = q.corners[2].x - q.corners[3].x;
    expect(topW).toBeLessThan(botW);
    expect(q.corners.map((c) => r(c, 2))).toMatchSnapshot();
  });
});

describe('computeDepthOffset (extrusion side-wall direction, §20.1.5.12)', () => {
  it('is ~zero for a face-on camera (−Z projects straight into the screen)', () => {
    const cam: CameraInput = { prst: 'perspectiveFront' };
    const o = computeDepthOffset(cam, W, H, 20);
    expect(Math.hypot(o.x, o.y)).toBeLessThan(0.5);
  });

  it('is exactly zero for orthographicFront (parallel, no tilt)', () => {
    const cam: CameraInput = { prst: 'orthographicFront' };
    const o = computeDepthOffset(cam, W, H, 20);
    expect(o.x).toBeCloseTo(0, 6);
    expect(o.y).toBeCloseTo(0, 6);
  });

  it('reveals a side wall when the camera is tilted (non-zero offset)', () => {
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 0 } };
    const o = computeDepthOffset(cam, W, H, 40);
    expect(Math.hypot(o.x, o.y)).toBeGreaterThan(1);
  });

  it('scales with the extrusion depth', () => {
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 0 } };
    const a = computeDepthOffset(cam, W, H, 20);
    const b = computeDepthOffset(cam, W, H, 40);
    expect(Math.hypot(b.x, b.y)).toBeGreaterThan(Math.hypot(a.x, a.y));
  });
});
