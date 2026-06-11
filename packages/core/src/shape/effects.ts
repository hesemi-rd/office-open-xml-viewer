import type { Shadow, SoftEdge, Reflection } from '../types/common';
import { hexToRgba } from './paint';

/**
 * Canvas 2D rendering of the three DrawingML edge/blur effects that the
 * Canvas shadow primitive (a single offset+blur slot) cannot express:
 *
 *   - innerShdw  — ECMA-376 §20.1.8.40 (CT_InnerShadowEffect)
 *   - softEdge   — ECMA-376 §20.1.8.53 (CT_SoftEdgesEffect)
 *   - reflection — ECMA-376 §20.1.8.50 (CT_ReflectionEffect)
 *
 * Each helper takes a `paintShape` callback that renders the shape's opaque
 * silhouette (fill + stroke) into a supplied 2D context, positioned in the
 * SAME device pixel coordinates as the live canvas. The helpers route that
 * silhouette through one or more auxiliary canvases and `globalCompositeOperation`
 * / `filter` passes, then blit the result back onto the live context — so the
 * caller never has to special-case preset vs. custom geometry.
 *
 * Coordinates: `bbox` is the shape bounding box in device pixels (already scaled).
 * Effect radii / distances arrive in EMU and are converted with `scale`.
 */

export type PaintShape = (ctx: AuxContext) => void;

type AuxCanvas = HTMLCanvasElement | OffscreenCanvas;
type AuxContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** Bounding box of the shape in device pixels (post-scale). */
export interface EffectBBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * EMU → device px. Mirrors `emuToPx` in the pptx renderer, where `scale` is
 * px-per-EMU (it already folds in the EMU→px conversion), so this is a bare
 * multiply. Kept local so the effect helpers have no dependency on the renderer.
 */
function emuToPx(emu: number, scale: number): number {
  return emu * scale;
}

/**
 * Allocate an auxiliary canvas. Prefers OffscreenCanvas (no DOM pollution, works
 * inside a worker); falls back to a detached <canvas>; returns null when neither
 * is available (e.g. a headless unit-test environment) so callers can no-op.
 * Dimensions are clamped to >=1 to avoid zero-size canvas errors.
 */
export function createAuxCanvas(w: number, h: number): AuxCanvas | null {
  const cw = Math.max(1, Math.ceil(w));
  const ch = Math.max(1, Math.ceil(h));
  if (typeof OffscreenCanvas !== 'undefined') {
    return new OffscreenCanvas(cw, ch);
  }
  if (typeof document !== 'undefined') {
    const c = document.createElement('canvas');
    c.width = cw;
    c.height = ch;
    return c;
  }
  return null;
}

function get2d(canvas: AuxCanvas): AuxContext | null {
  // The union return of getContext('2d') across HTMLCanvasElement /
  // OffscreenCanvas is awkward to type; narrow via unknown.
  return (canvas.getContext('2d') as AuxContext | null) ?? null;
}

/**
 * innerShdw — ECMA-376 §20.1.8.40. "A shadow is applied within the edges of the
 * object." The shadow colour shows through where the shape's interior is NOT
 * covered by an offset copy of the shape — i.e. it hugs the inside of the edge
 * opposite the offset direction.
 *
 * Construction (all on an auxiliary canvas the size of the live canvas so the
 * shadow can bleed up to `blur` px past the silhouette before being clipped):
 *   1. Paint the silhouette filled with the shadow colour.
 *   2. `destination-out` an offset + blurred copy of the silhouette. What
 *      remains is a blurred crescent inside the original edge — the inner shadow.
 *   3. `destination-in` the (un-offset, un-blurred) silhouette so the shadow is
 *      clipped to the shape interior.
 *   4. Blit over the live shape with `source-over` (drawn AFTER the fill).
 *
 * Offset uses dir (degrees clockwise from East) and dist (EMU). The shadow is
 * cast toward `dir`, so the colour collects on the far side — we punch the
 * offset copy toward `dir` and the surviving colour sits opposite it, matching
 * PowerPoint's inset shadow.
 */
export function applyInnerShadow(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  _bbox: EffectBBox,
  shadow: Shadow,
  scale: number,
  deviceW: number,
  deviceH: number,
): void {
  const aux = createAuxCanvas(deviceW, deviceH);
  if (!aux) return;
  const c = get2d(aux);
  if (!c) return;

  const blur = emuToPx(shadow.blur, scale);
  const dist = emuToPx(shadow.dist, scale);
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dx = Math.cos(dirRad) * dist;
  const dy = Math.sin(dirRad) * dist;

  // 1. Silhouette in the shadow colour.
  c.save();
  c.fillStyle = hexToRgba(shadow.color, shadow.alpha);
  paintShape(c);
  c.restore();

  // 2. Carve out the offset + blurred copy; the leftover is the inner shadow.
  c.save();
  c.globalCompositeOperation = 'destination-out';
  c.filter = blur > 0 ? `blur(${blur}px)` : 'none';
  c.translate(dx, dy);
  c.fillStyle = '#000';
  paintShape(c);
  c.restore();

  // 3. Clip the surviving shadow to the shape interior.
  c.save();
  c.globalCompositeOperation = 'destination-in';
  c.filter = 'none';
  c.fillStyle = '#000';
  paintShape(c);
  c.restore();

  // 4. Composite over the live shape.
  liveCtx.save();
  liveCtx.drawImage(aux as CanvasImageSource, 0, 0);
  liveCtx.restore();
}

/**
 * softEdge — ECMA-376 §20.1.8.53. "The edges of the shape are blurred, while the
 * fill is not affected." We feather the shape's alpha by `rad` EMU: build a
 * blurred silhouette mask and `destination-in` it onto the already-painted shape
 * so the perimeter fades to transparent over the feather radius.
 *
 * This helper paints the shape onto an auxiliary canvas, then intersects it with
 * a blurred version of its FILLED SILHOUETTE (not the styled body), then blits
 * the feathered result back. The mask must be the filled silhouette — masking
 * with a stroked body would `destination-in` only the thin outline ring and
 * erase the interior. Because the silhouette interior stays alpha = 1 under
 * blur, only the edge band fades, matching the spec's "fill is not affected".
 *
 * `paintMask` paints a flat opaque silhouette (filled path, no stroke). When
 * omitted, `paintShape` is reused (correct only for unstroked shapes).
 */
export function applySoftEdge(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  _bbox: EffectBBox,
  softEdge: SoftEdge,
  scale: number,
  deviceW: number,
  deviceH: number,
  paintMask?: PaintShape,
): void {
  const radius = emuToPx(softEdge.radius, scale);
  if (radius <= 0) {
    // No feather: paint straight onto the live context.
    paintShape(liveCtx);
    return;
  }

  const aux = createAuxCanvas(deviceW, deviceH);
  if (!aux) {
    paintShape(liveCtx);
    return;
  }
  const c = get2d(aux);
  if (!c) {
    paintShape(liveCtx);
    return;
  }

  const mask = paintMask ?? paintShape;

  // 1. Paint the full shape (fill + stroke) onto the aux canvas.
  paintShape(c);

  // 2. Intersect with a blurred FILLED silhouette to feather the perimeter.
  //    ECMA gives `rad` as the blur radius directly.
  c.save();
  c.globalCompositeOperation = 'destination-in';
  c.filter = `blur(${radius}px)`;
  c.fillStyle = '#000';
  mask(c);
  c.restore();

  // 3. Blit the feathered shape back to the live context.
  liveCtx.save();
  liveCtx.drawImage(aux as CanvasImageSource, 0, 0);
  liveCtx.restore();
}

/**
 * reflection — ECMA-376 §20.1.8.50. A mirrored copy of the shape rendered below
 * it, faded out by a linear alpha gradient.
 *
 * Steps:
 *   1. Paint the shape onto an aux canvas (its own silhouette + fill/stroke).
 *   2. Apply a vertical alpha gradient via `destination-in`: alpha runs from
 *      `stA` at `stPos` to `endA` at `endPos` along the (post-mirror top→bottom)
 *      ramp. With the default fadeDir=5400000 (90°, downward) the ramp is the
 *      reflection's own vertical axis.
 *   3. Blit the faded copy onto the live context under a transform that mirrors
 *      it (sx/sy), offsets it by `dist` along `dir`, and anchors it to the
 *      shape's bottom edge (algn="b" default).
 *
 * Skew (kx/ky), non-bottom alignment, fadeDir != 90°, and rotWithShape are not
 * carried by the parser model; their spec defaults are assumed (kx=ky=0,
 * algn="b", fadeDir=90°, rotWithShape=true → reflection shares the shape's
 * rotation, which the caller's transform already provides).
 */
export function applyReflection(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  bbox: EffectBBox,
  reflection: Reflection,
  scale: number,
  deviceW: number,
  deviceH: number,
): void {
  const aux = createAuxCanvas(deviceW, deviceH);
  if (!aux) return;
  const c = get2d(aux);
  if (!c) return;

  const blur = emuToPx(reflection.blur, scale);

  // 1. Paint the shape onto the aux canvas, optionally blurred.
  c.save();
  if (blur > 0) c.filter = `blur(${blur}px)`;
  paintShape(c);
  c.restore();

  // 2. Fade with a vertical alpha gradient over the shape's bbox band.
  //    Map stPos/endPos (0..1 along the ramp) to the bbox top..bottom.
  c.save();
  c.globalCompositeOperation = 'destination-in';
  const top = bbox.y;
  const bottom = bbox.y + bbox.h;
  const grad = c.createLinearGradient(0, top, 0, bottom);
  const stPos = clamp01(reflection.stPos);
  const endPos = clamp01(reflection.endPos);
  // Guard equal stops (createLinearGradient requires strictly increasing offsets
  // only loosely, but identical offsets collapse the ramp). stA applies before
  // stPos, endA after endPos — pad the ends so the band outside [stPos,endPos]
  // holds the terminal alpha.
  grad.addColorStop(0, `rgba(0,0,0,${reflection.stA})`);
  if (stPos > 0) grad.addColorStop(stPos, `rgba(0,0,0,${reflection.stA})`);
  if (endPos < 1 && endPos > stPos) grad.addColorStop(endPos, `rgba(0,0,0,${reflection.endA})`);
  grad.addColorStop(1, `rgba(0,0,0,${reflection.endA})`);
  c.fillStyle = grad;
  c.fillRect(0, 0, deviceW, deviceH);
  c.restore();

  // 3. Blit the faded mirror under the shape. Mirror about the shape's bottom
  //    edge (algn="b"): reflect across y = bottom, then push down by `dist`.
  const dist = emuToPx(reflection.dist, scale);
  const dirRad = (reflection.dir * Math.PI) / 180;
  const offX = Math.cos(dirRad) * dist;
  const offY = Math.sin(dirRad) * dist;

  liveCtx.save();
  // Translate to the bottom edge, apply sx/sy scale (sy<0 = mirror), then
  // translate back. Anchoring at the bottom edge keeps the reflection's top
  // touching the shape's bottom before `dist` separation.
  liveCtx.translate(bbox.x + offX, bottom + offY);
  liveCtx.scale(reflection.sx, reflection.sy);
  liveCtx.translate(-(bbox.x), -bottom);
  liveCtx.drawImage(aux as CanvasImageSource, 0, 0);
  liveCtx.restore();
}

function clamp01(v: number): number {
  return v < 0 ? 0 : v > 1 ? 1 : v;
}
