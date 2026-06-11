/**
 * Draw a 2D bitmap warped into a projected quad (DrawingML 3D camera, Phase A).
 *
 * Canvas 2D has no native perspective (homography) transform — `setTransform`
 * is affine only. We therefore map the source rectangle onto the destination
 * quad with a piecewise-affine mesh: the source is split into a grid of cells,
 * each cell is drawn with its own affine `setTransform` derived from the
 * homography, and the grid is refined until the affine approximation error of
 * every cell is below a sub-pixel threshold.
 *
 * This works with plain Canvas 2D (and skia-canvas in packages/node) — no
 * WebGL — per the project's renderer constraints.
 */

import type { Vec2 } from './scene3d-camera';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
type AnyImage = CanvasImageSource;

/** 3×3 homography in row-major order. */
type H = [number, number, number, number, number, number, number, number, number];

/**
 * Solve the homography H that maps the unit square corners
 *   (0,0) (1,0) (1,1) (0,1)
 * to the four destination points d0..d3 (same TL,TR,BR,BL order).
 *
 * Standard projective mapping of a unit square to a quadrilateral
 * (Heckbert, "Fundamentals of Texture Mapping and Image Warping", §2.2).
 * Returns null if the quad is degenerate.
 */
function unitSquareToQuad(d0: Vec2, d1: Vec2, d2: Vec2, d3: Vec2): H | null {
  const x0 = d0.x,
    y0 = d0.y;
  const x1 = d1.x,
    y1 = d1.y;
  const x2 = d2.x,
    y2 = d2.y;
  const x3 = d3.x,
    y3 = d3.y;

  const dx1 = x1 - x2;
  const dx2 = x3 - x2;
  const dx3 = x0 - x1 + x2 - x3;
  const dy1 = y1 - y2;
  const dy2 = y3 - y2;
  const dy3 = y0 - y1 + y2 - y3;

  let a13: number;
  let a23: number;
  if (Math.abs(dx3) < 1e-12 && Math.abs(dy3) < 1e-12) {
    // Affine (parallelogram) case: g = h = 0.
    a13 = 0;
    a23 = 0;
  } else {
    const den = dx1 * dy2 - dx2 * dy1;
    if (Math.abs(den) < 1e-12) return null;
    a13 = (dx3 * dy2 - dx2 * dy3) / den;
    a23 = (dx1 * dy3 - dx3 * dy1) / den;
  }
  const a11 = x1 - x0 + a13 * x1;
  const a21 = x3 - x0 + a23 * x3;
  const a31 = x0;
  const a12 = y1 - y0 + a13 * y1;
  const a22 = y3 - y0 + a23 * y3;
  const a32 = y0;
  // Row-major: [a11 a21 a31; a12 a22 a32; a13 a23 1]
  return [a11, a21, a31, a12, a22, a32, a13, a23, 1];
}

/** Apply homography H to a unit-square coordinate (u,v) ∈ [0,1]². */
function applyH(h: H, u: number, v: number): Vec2 {
  const w = h[6] * u + h[7] * v + h[8];
  return {
    x: (h[0] * u + h[1] * v + h[2]) / w,
    y: (h[3] * u + h[4] * v + h[5]) / w,
  };
}

/**
 * Draw the affine image of one source cell. Given the source-space rectangle
 * [sx0,sx1]×[sy0,sy1] and the destination positions of its three corners
 * (origin p0=TL, pu=TR, pv=BL), compute the affine matrix that sends the source
 * rect to that parallelogram and `drawImage` the sub-rectangle through it.
 *
 * A 0.5px outward bleed on the source sub-rect (and a matching dest expansion)
 * hides the hairline seams that otherwise appear between mesh cells from
 * fractional-pixel sampling. The bleed is symmetric so adjacent cells overlap
 * rather than gap.
 */
function drawCell(
  ctx: AnyCtx,
  img: AnyImage,
  imgW: number,
  imgH: number,
  sx0: number,
  sy0: number,
  sx1: number,
  sy1: number,
  p0: Vec2,
  pu: Vec2,
  pv: Vec2,
): void {
  const sw = sx1 - sx0;
  const sh = sy1 - sy0;
  if (sw <= 0 || sh <= 0) return;

  // Destination basis vectors (per unit of source u / v).
  const ux = (pu.x - p0.x) / sw;
  const uy = (pu.y - p0.y) / sw;
  const vx = (pv.x - p0.x) / sh;
  const vy = (pv.y - p0.y) / sh;

  // Seam-hiding bleed: half a device pixel, expressed in source units along
  // each axis (so the dest overlap is ~0.5px regardless of the cell's scale).
  const destLenU = Math.hypot(pu.x - p0.x, pu.y - p0.y) || 1;
  const destLenV = Math.hypot(pv.x - p0.x, pv.y - p0.y) || 1;
  const bleedU = (0.5 * sw) / destLenU;
  const bleedV = (0.5 * sh) / destLenV;

  // Clamp the bled source rect into the image so we never sample outside it.
  const bx0 = Math.max(0, sx0 - bleedU);
  const by0 = Math.max(0, sy0 - bleedV);
  const bx1 = Math.min(imgW, sx1 + bleedU);
  const by1 = Math.min(imgH, sy1 + bleedV);
  const bw = bx1 - bx0;
  const bh = by1 - by0;
  if (bw <= 0 || bh <= 0) return;

  ctx.save();
  // setTransform maps source pixel (sx,sy) → dest:
  //   dest = p0 + (sx - sx0)*u_basis + (sy - sy0)*v_basis
  // i.e. matrix (a,b,c,d,e,f) with a=ux, b=uy, c=vx, d=vy and translation
  // chosen so source (sx0,sy0) lands on p0.
  const e = p0.x - sx0 * ux - sy0 * vx;
  const f = p0.y - sx0 * uy - sy0 * vy;
  ctx.setTransform(ux, uy, vx, vy, e, f);
  ctx.drawImage(img, bx0, by0, bw, bh, bx0, by0, bw, bh);
  ctx.restore();
}

/**
 * Recursively warp source sub-rect [u0,u1]×[v0,v1] (unit-square coords) into the
 * destination quad described by homography H. A cell is drawn directly when its
 * centre's affine-vs-homography disagreement is below `tol` device px; otherwise
 * it is split along its longer axis and recursed. `depth` caps recursion so a
 * pathological quad can't blow the stack.
 */
function warpRecursive(
  ctx: AnyCtx,
  img: AnyImage,
  imgW: number,
  imgH: number,
  h: H,
  u0: number,
  v0: number,
  u1: number,
  v1: number,
  tol: number,
  depth: number,
): void {
  // Destination corners of this cell.
  const c00 = applyH(h, u0, v0); // TL
  const c10 = applyH(h, u1, v0); // TR
  const c01 = applyH(h, u0, v1); // BL
  const c11 = applyH(h, u1, v1); // BR

  // Affine prediction of the cell centre (bilinear midpoint of the 4 dest
  // corners) vs. the true homography position of the centre. Their distance is
  // the perspective error this cell would incur if drawn as a single affine
  // parallelogram.
  const um = (u0 + u1) / 2;
  const vm = (v0 + v1) / 2;
  const trueMid = applyH(h, um, vm);
  const affMid = {
    x: (c00.x + c10.x + c01.x + c11.x) / 4,
    y: (c00.y + c10.y + c01.y + c11.y) / 4,
  };
  const err = Math.hypot(trueMid.x - affMid.x, trueMid.y - affMid.y);

  if (depth <= 0 || err <= tol) {
    // Draw as two affine triangles' worth via a single parallelogram derived
    // from the TL/TR/BL corners. The mesh is fine enough that the BR corner's
    // residual is within tol, so one affine cell suffices.
    const sx0 = u0 * imgW;
    const sy0 = v0 * imgH;
    const sx1 = u1 * imgW;
    const sy1 = v1 * imgH;
    drawCell(ctx, img, imgW, imgH, sx0, sy0, sx1, sy1, c00, c10, c01);
    return;
  }

  // Split along the longer source axis to keep cells roughly square.
  const du = u1 - u0;
  const dv = v1 - v0;
  if (du >= dv) {
    warpRecursive(ctx, img, imgW, imgH, h, u0, v0, um, v1, tol, depth - 1);
    warpRecursive(ctx, img, imgW, imgH, h, um, v0, u1, v1, tol, depth - 1);
  } else {
    warpRecursive(ctx, img, imgW, imgH, h, u0, v0, u1, vm, tol, depth - 1);
    warpRecursive(ctx, img, imgW, imgH, h, u0, vm, u1, v1, tol, depth - 1);
  }
}

/**
 * Warp `src` (a `srcW`×`srcH` bitmap, e.g. an offscreen canvas of the shape's
 * normal 2D rendering) into the destination quad `corners` on `dst`.
 *
 * @param corners  destination quad in TL,TR,BR,BL order, in `dst` pixel space.
 *                 These come from `computeScene3dQuad`, already offset to the
 *                 element's position by the caller.
 * @param tol      max affine-approximation error per cell, in device px. The
 *                 mesh subdivides until every cell is within this. Default 0.5px
 *                 (sub-pixel — invisible at 1× and most HiDPI ratios). Not a
 *                 magic cell count: the subdivision is error-driven.
 */
export function drawProjected(
  src: AnyImage,
  dst: AnyCtx,
  srcW: number,
  srcH: number,
  corners: [Vec2, Vec2, Vec2, Vec2],
  tol = 0.5,
): void {
  if (srcW <= 0 || srcH <= 0) return;
  // Reject a degenerate (near-zero-area) destination quad. The signed area of
  // the polygon TL,TR,BR,BL via the shoelace formula; |area| ~ 0 means the quad
  // collapsed to a line / point and there is nothing to draw.
  const [p0, p1, p2, p3] = corners;
  const area =
    Math.abs(
      p0.x * p1.y - p1.x * p0.y +
        p1.x * p2.y - p2.x * p1.y +
        p2.x * p3.y - p3.x * p2.y +
        p3.x * p0.y - p0.x * p3.y,
    ) / 2;
  if (area < 1e-6) return;
  const h = unitSquareToQuad(corners[0], corners[1], corners[2], corners[3]);
  if (!h) return; // degenerate quad — skip rather than draw garbage.

  // Recursion cap: 2^7 = 128 splits per axis is far past sub-pixel for any sane
  // tilt; the error test almost always stops much sooner.
  const MAX_DEPTH = 14;
  dst.save();
  // Clip to the quad so the bleed overlap can't spill past the projected shape.
  dst.beginPath();
  dst.moveTo(corners[0].x, corners[0].y);
  dst.lineTo(corners[1].x, corners[1].y);
  dst.lineTo(corners[2].x, corners[2].y);
  dst.lineTo(corners[3].x, corners[3].y);
  dst.closePath();
  dst.clip();
  warpRecursive(dst, src, srcW, srcH, h, 0, 0, 1, 1, tol, MAX_DEPTH);
  dst.restore();
}
