import { describe, it, expect, beforeAll, afterAll } from 'vitest';
// Deep-import the effect helpers straight from the core SOURCE (mirrors how
// bevel-shading.test.ts deep-imports its module): `@silurus/ooxml-core` is a
// transitive dep here (node → pptx → core) but not symlinked into this package's
// node_modules, so a relative source path is the reliable resolution for vitest.
import {
  applyInnerShadow,
  applySoftEdge,
  applyReflection,
  type EffectBBox,
  type PaintShape,
} from '../../core/src/shape/effects';
import type { Shadow, SoftEdge, Reflection } from '../../core/src/types/common';
import { installOffscreenCanvasShim, type NodeCanvasFactory } from './render';
import { loadSkiaForTests } from './test-imports';

/**
 * A4 pixel-equality oracle — cropped aux canvases vs the full-canvas version.
 *
 * `packages/core/src/shape/effects.ts` was changed (perf: A4) to allocate its
 * auxiliary canvases at the shape's device bbox + effect margin instead of the
 * whole live canvas. The claim is that the on-screen pixels are UNCHANGED, one
 * channel at a time. This suite proves it on real skia-canvas pixels:
 *
 *   1. Verbatim FULL-CANVAS oracles (`oracleInnerShadow` / `oracleSoftEdge` /
 *      `oracleReflection`) reproduce the pre-A4 code exactly — aux canvas =
 *      deviceW × deviceH, blit at (0,0).
 *   2. The SAME shape + effect is rendered twice into two identical live skia
 *      canvases: once by the real (cropped) core helper, once by the oracle.
 *   3. Every pixel must match to the byte (`expectExactPixels`).
 *
 * The paint callback mimics the pptx renderer: it `setTransform(liveTransform)`
 * to place the silhouette in absolute device pixels, exactly the callback shape
 * the crop's transform-offset proxy must transparently handle. A shape placed
 * OFF-ORIGIN (and, in one case, against the canvas edge to exercise clamping)
 * makes the crop non-trivial, so an off-by-one in the offset or blit would show.
 */

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, DOMMatrix } = (skia ?? {}) as Skia;

type Ctx2D = import('skia-canvas').CanvasRenderingContext2D;

const DEVICE_W = 320;
const DEVICE_H = 240;
const SCALE = 1 / 9525; // px-per-EMU; 9525 EMU → 1 device px (mirrors the pptx renderer)

/** The absolute device transform the renderer installs (here: dpr=2, no rot).
 *  skia-canvas ships its own DOMMatrix (Node has no global one); a plain 2× scale
 *  is enough to exercise the crop proxy — the shape's device coords differ from
 *  its CSS coords, so the offset composition into `setTransform` must be exact. */
function liveTransform(): import('skia-canvas').DOMMatrix {
  return new DOMMatrix().scaleSelf(2, 2);
}

/**
 * Build a paint callback that draws a rounded shape (fill + stroke) at CSS-space
 * (cx,cy,cw,ch), applying the live transform first — exactly as the renderer's
 * `(c) => { applyLiveTransform(c); paintShapeBody(c); }` does. The silhouette
 * variant (a flat opaque fill, no stroke) is used for masks.
 */
function makePaint(cx: number, cy: number, cw: number, ch: number): {
  paint: PaintShape;
  mask: PaintShape;
  /** Device-space bbox (CSS bbox × the 2× live transform). */
  bbox: EffectBBox;
} {
  const draw = (c: Ctx2D, silhouette?: string): void => {
    c.setTransform(liveTransform());
    c.beginPath();
    // A shape with curved + straight edges so the blur/feather has real gradients.
    const r = Math.min(cw, ch) * 0.25;
    c.moveTo(cx + r, cy);
    c.lineTo(cx + cw - r, cy);
    c.quadraticCurveTo(cx + cw, cy, cx + cw, cy + r);
    c.lineTo(cx + cw, cy + ch - r);
    c.quadraticCurveTo(cx + cw, cy + ch, cx + cw - r, cy + ch);
    c.lineTo(cx + r, cy + ch);
    c.quadraticCurveTo(cx, cy + ch, cx, cy + ch - r);
    c.lineTo(cx, cy + r);
    c.quadraticCurveTo(cx, cy, cx + r, cy);
    c.closePath();
    c.fillStyle = silhouette ?? '#4488cc';
    c.fill();
    if (!silhouette) {
      c.lineWidth = 2;
      c.strokeStyle = '#113355';
      c.stroke();
    }
  };
  return {
    paint: (c) => draw(c as unknown as Ctx2D),
    mask: (c) => draw(c as unknown as Ctx2D, '#000'),
    bbox: { x: cx * 2, y: cy * 2, w: cw * 2, h: ch * 2 },
  };
}

// ── Verbatim full-canvas oracles (pre-A4 effects.ts) ────────────────────────

function hexToRgbaLocal(hex: string, alpha: number): string {
  const h = hex.replace('#', '');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}
function newAux(w: number, h: number): Ctx2D {
  const c = new Canvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h)));
  return c.getContext('2d') as unknown as Ctx2D;
}

function oracleInnerShadow(
  liveCtx: Ctx2D, paintShape: PaintShape, shadow: Shadow, scale: number, dW: number, dH: number,
): void {
  const c = newAux(dW, dH);
  const blur = shadow.blur * scale;
  const dist = shadow.dist * scale;
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dx = Math.cos(dirRad) * dist;
  const dy = Math.sin(dirRad) * dist;
  c.save();
  c.fillStyle = hexToRgbaLocal(shadow.color, shadow.alpha);
  paintShape(c as never);
  c.restore();
  c.save();
  c.globalCompositeOperation = 'destination-out';
  c.filter = blur > 0 ? `blur(${blur}px)` : 'none';
  c.translate(dx, dy);
  c.fillStyle = '#000';
  paintShape(c as never);
  c.restore();
  c.save();
  c.globalCompositeOperation = 'destination-in';
  c.filter = 'none';
  c.fillStyle = '#000';
  paintShape(c as never);
  c.restore();
  liveCtx.save();
  liveCtx.drawImage(c.canvas as never, 0, 0);
  liveCtx.restore();
}

function oracleSoftEdge(
  liveCtx: Ctx2D, paintShape: PaintShape, bbox: EffectBBox, se: SoftEdge,
  scale: number, dW: number, dH: number, paintMask: PaintShape,
): void {
  const radius = se.radius * scale;
  if (radius <= 0) { paintShape(liveCtx as never); return; }
  const c = newAux(dW, dH);
  const mc = newAux(dW, dH);
  const cc = newAux(dW, dH);
  paintShape(c as never);
  mc.fillStyle = '#000';
  paintMask(mc as never);
  cc.drawImage(
    c.canvas as never,
    bbox.x, bbox.y, bbox.w, bbox.h,
    bbox.x - radius, bbox.y - radius, bbox.w + radius * 2, bbox.h + radius * 2,
  );
  cc.drawImage(c.canvas as never, 0, 0);
  cc.globalCompositeOperation = 'destination-in';
  cc.filter = `blur(${radius / 3}px)`;
  cc.drawImage(mc.canvas as never, 0, 0);
  cc.filter = 'none';
  cc.globalCompositeOperation = 'source-over';
  liveCtx.save();
  liveCtx.drawImage(cc.canvas as never, 0, 0);
  liveCtx.restore();
}

function oracleReflection(
  liveCtx: Ctx2D, paintShape: PaintShape, bbox: EffectBBox, reflection: Reflection,
  scale: number, dW: number, dH: number,
): void {
  const c = newAux(dW, dH);
  const blur = reflection.blur * scale;
  c.save();
  if (blur > 0) c.filter = `blur(${blur}px)`;
  paintShape(c as never);
  c.restore();
  c.save();
  c.globalCompositeOperation = 'destination-in';
  const top = bbox.y;
  const bottom = bbox.y + bbox.h;
  const grad = c.createLinearGradient(0, bottom, 0, top);
  const clamp01 = (v: number) => (v < 0 ? 0 : v > 1 ? 1 : v);
  const stPos = clamp01(reflection.stPos);
  const endPos = clamp01(reflection.endPos);
  grad.addColorStop(0, `rgba(0,0,0,${reflection.stA})`);
  if (stPos > 0) grad.addColorStop(stPos, `rgba(0,0,0,${reflection.stA})`);
  if (endPos < 1 && endPos > stPos) grad.addColorStop(endPos, `rgba(0,0,0,${reflection.endA})`);
  grad.addColorStop(1, `rgba(0,0,0,${reflection.endA})`);
  c.fillStyle = grad;
  c.fillRect(0, 0, dW, dH);
  c.restore();
  const dist = reflection.dist * scale;
  const dirRad = (reflection.dir * Math.PI) / 180;
  const offX = Math.cos(dirRad) * dist;
  const offY = Math.sin(dirRad) * dist;
  liveCtx.save();
  liveCtx.translate(bbox.x + offX, bottom + offY);
  liveCtx.scale(reflection.sx, reflection.sy);
  liveCtx.translate(-bbox.x, -bottom);
  liveCtx.drawImage(c.canvas as never, 0, 0);
  liveCtx.restore();
}

// ── Comparison harness ──────────────────────────────────────────────────────

function freshLive(): { canvas: import('skia-canvas').Canvas; ctx: Ctx2D } {
  const canvas = new Canvas(DEVICE_W, DEVICE_H);
  const ctx = canvas.getContext('2d') as unknown as Ctx2D;
  // Fill with an opaque backdrop so any compositing difference (alpha handling)
  // shows up as an RGB difference, not just an alpha one.
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, DEVICE_W, DEVICE_H);
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  return { canvas, ctx };
}

function pixels(ctx: Ctx2D): Uint8ClampedArray {
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  return ctx.getImageData(0, 0, DEVICE_W, DEVICE_H).data as unknown as Uint8ClampedArray;
}

/** Assert two RGBA buffers are byte-identical; report the worst pixel if not. */
function expectExactPixels(a: Uint8ClampedArray, b: Uint8ClampedArray, label: string): void {
  expect(a.length).toBe(b.length);
  let diffPixels = 0;
  let maxDelta = 0;
  let worstAt = -1;
  for (let i = 0; i < a.length; i += 4) {
    let d = 0;
    for (let k = 0; k < 4; k++) d = Math.max(d, Math.abs(a[i + k] - b[i + k]));
    if (d > 0) {
      diffPixels++;
      if (d > maxDelta) { maxDelta = d; worstAt = i; }
    }
  }
  if (diffPixels > 0) {
    const px = worstAt / 4;
    throw new Error(
      `${label}: ${diffPixels} pixel(s) differ, max channel delta ${maxDelta} ` +
        `at (${px % DEVICE_W},${Math.floor(px / DEVICE_W)}) ` +
        `cropped=[${a[worstAt]},${a[worstAt + 1]},${a[worstAt + 2]},${a[worstAt + 3]}] ` +
        `oracle=[${b[worstAt]},${b[worstAt + 1]},${b[worstAt + 2]},${b[worstAt + 3]}]`,
    );
  }
  expect(diffPixels).toBe(0);
}

// The real core `createAuxCanvas` allocates via `new OffscreenCanvas(w,h)`, which
// Node lacks. Install the same skia-backed shim `renderSlideNode` uses so the
// cropped helper's auxiliary canvases are real skia canvases; without it the
// helper returns early (null aux) and paints nothing — masking the comparison.
let restoreShim: (() => void) | null = null;
beforeAll(() => {
  if (!skia) return;
  const factory: NodeCanvasFactory = {
    createCanvas: (w, h) =>
      new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
    // loadImage is unused by these effect helpers (no picture decode), but the
    // factory type requires it.
    loadImage: (() => {
      throw new Error('loadImage not used in effect-crop tests');
    }) as unknown as NodeCanvasFactory['loadImage'],
  };
  restoreShim = installOffscreenCanvasShim(factory);
});
afterAll(() => {
  restoreShim?.();
  restoreShim = null;
});

describe.skipIf(!skia)('A4 effect crop — byte-exact vs full-canvas oracle (real skia pixels)', () => {
  // Shape placed WELL INSIDE the canvas (CSS 30,20 → device 60,40; the crop is a
  // small interior rectangle, the common case A4 optimises).
  const INTERIOR = makePaint(30, 20, 80, 50);
  // Shape hugging the top-left ORIGIN so the crop clamps at x=0/y=0 (exercises the
  // `crop.x===0 && crop.y===0` proxy short-circuit + the blit at the clamped edge).
  const AT_ORIGIN = makePaint(1, 1, 70, 45);

  it('innerShadow: cropped output equals full-canvas, interior shape', () => {
    const shadow: Shadow = { color: '303030', alpha: 0.75, blur: 28575, dist: 19050, dir: 135 };
    const real = freshLive();
    applyInnerShadow(real.ctx as never, INTERIOR.paint, INTERIOR.bbox, shadow, SCALE, DEVICE_W, DEVICE_H);
    const oracle = freshLive();
    oracleInnerShadow(oracle.ctx, INTERIOR.paint, shadow, SCALE, DEVICE_W, DEVICE_H);
    expectExactPixels(pixels(real.ctx), pixels(oracle.ctx), 'innerShadow interior');
  });

  it('innerShadow: cropped output equals full-canvas, shape clamped at origin', () => {
    const shadow: Shadow = { color: '000000', alpha: 0.6, blur: 19050, dist: 9525, dir: 45 };
    const real = freshLive();
    applyInnerShadow(real.ctx as never, AT_ORIGIN.paint, AT_ORIGIN.bbox, shadow, SCALE, DEVICE_W, DEVICE_H);
    const oracle = freshLive();
    oracleInnerShadow(oracle.ctx, AT_ORIGIN.paint, shadow, SCALE, DEVICE_W, DEVICE_H);
    expectExactPixels(pixels(real.ctx), pixels(oracle.ctx), 'innerShadow origin');
  });

  it('softEdge: cropped output equals full-canvas, interior shape', () => {
    const se: SoftEdge = { radius: 47625 }; // 5 px feather
    const real = freshLive();
    applySoftEdge(real.ctx as never, INTERIOR.paint, INTERIOR.bbox, se, SCALE, DEVICE_W, DEVICE_H, INTERIOR.mask);
    const oracle = freshLive();
    oracleSoftEdge(oracle.ctx, INTERIOR.paint, INTERIOR.bbox, se, SCALE, DEVICE_W, DEVICE_H, INTERIOR.mask);
    expectExactPixels(pixels(real.ctx), pixels(oracle.ctx), 'softEdge interior');
  });

  it('softEdge: cropped output equals full-canvas, shape clamped at origin', () => {
    const se: SoftEdge = { radius: 28575 }; // 3 px feather
    const real = freshLive();
    applySoftEdge(real.ctx as never, AT_ORIGIN.paint, AT_ORIGIN.bbox, se, SCALE, DEVICE_W, DEVICE_H, AT_ORIGIN.mask);
    const oracle = freshLive();
    oracleSoftEdge(oracle.ctx, AT_ORIGIN.paint, AT_ORIGIN.bbox, se, SCALE, DEVICE_W, DEVICE_H, AT_ORIGIN.mask);
    expectExactPixels(pixels(real.ctx), pixels(oracle.ctx), 'softEdge origin');
  });

  it('reflection: cropped output equals full-canvas, interior shape', () => {
    const reflection: Reflection = {
      blur: 19050, dist: 9525, dir: 90, stA: 0.6, stPos: 0, endA: 0, endPos: 0.4, sx: 1, sy: -1,
    };
    const real = freshLive();
    applyReflection(real.ctx as never, INTERIOR.paint, INTERIOR.bbox, reflection, SCALE, DEVICE_W, DEVICE_H);
    const oracle = freshLive();
    oracleReflection(oracle.ctx, INTERIOR.paint, INTERIOR.bbox, reflection, SCALE, DEVICE_W, DEVICE_H);
    expectExactPixels(pixels(real.ctx), pixels(oracle.ctx), 'reflection interior');
  });

  it('reflection: cropped output equals full-canvas with no blur (sharp mirror)', () => {
    const reflection: Reflection = {
      blur: 0, dist: 4762, dir: 90, stA: 0.5, stPos: 0.1, endA: 0.05, endPos: 0.5, sx: 1, sy: -1,
    };
    const real = freshLive();
    applyReflection(real.ctx as never, INTERIOR.paint, INTERIOR.bbox, reflection, SCALE, DEVICE_W, DEVICE_H);
    const oracle = freshLive();
    oracleReflection(oracle.ctx, INTERIOR.paint, INTERIOR.bbox, reflection, SCALE, DEVICE_W, DEVICE_H);
    expectExactPixels(pixels(real.ctx), pixels(oracle.ctx), 'reflection no-blur');
  });
});
