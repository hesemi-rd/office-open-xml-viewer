// ── Canvas backing-store size clamp (browser-limit guard) ────────────────────
//
// Every browser silently caps the pixel dimensions of a `<canvas>` /
// `OffscreenCanvas` backing store. Assigning `canvas.width`/`.height` beyond the
// cap does NOT throw — the browser allocates a SMALLER buffer (or, past the
// area limit, a zero-size one), so the canvas renders blank or truncated with no
// error. A document with a pathological page/slide size, or a very large
// devicePixelRatio × page size, can trip this and produce a blank viewer.
//
// `clampCanvasSize` bounds a requested backing size to limits that every major
// browser honors, scaling BOTH axes by the same factor so the aspect ratio is
// preserved (a uniformly smaller canvas that the caller then stretches to the
// intended CSS box, rather than a distorted or blank one).

/**
 * Maximum canvas backing-store dimension (px) per axis. 32767 is the largest a
 * `<canvas>` / `OffscreenCanvas` axis every major desktop browser accepts
 * (Chrome, Firefox and Safari all top out at 32767 on at least one axis; older
 * WebKit is lower but the area cap below binds first there).
 */
export const MAX_CANVAS_DIMENSION = 32767;

/**
 * Maximum canvas backing-store AREA (total px). Safari/WebKit is the tightest of
 * the major engines here: its canvas area limit is the smallest, so a canvas
 * well within the per-axis cap can still exceed the area cap and blank out. We
 * use 16,777,216 px (2^24 = 4096×4096-equivalent) — the widely-cited iOS/WebKit
 * canvas-area ceiling — as a conservative cross-engine bound. A page rendered at
 * or below this always has a real backing store on every engine; a larger
 * request is scaled down uniformly (and drawn back up to its CSS box).
 *
 * A4 at 96 DPI is ~0.6 MP and at a 3× DPR still ~5.3 MP — comfortably inside the
 * budget; the clamp only bites on genuinely pathological sizes.
 */
export const MAX_CANVAS_AREA = 1 << 24; // 16_777_216 px

/** A clamped canvas size plus the uniform scale applied (1 when unclamped). */
export interface ClampedCanvasSize {
  /** Clamped, integer backing-store width (≥ 1). */
  width: number;
  /** Clamped, integer backing-store height (≥ 1). */
  height: number;
  /**
   * The uniform factor applied to BOTH requested axes (`≤ 1`). 1 means no
   * clamping happened; `< 1` means the caller should draw the clamped canvas
   * scaled up by `1 / scale` to fill the originally-intended CSS box.
   */
  scale: number;
  /** True when the requested size exceeded a limit and was scaled down. */
  clamped: boolean;
}

/**
 * Clamp a requested canvas backing size to the per-axis and total-area limits,
 * preserving aspect ratio.
 *
 * The requested `width`/`height` are first rounded to integers and floored to
 * `≥ 1` (so a sub-0.5px request, which `Math.round` would send to 0, still
 * yields a real 1×1 backing store rather than a 0×0 one that renders blank). If
 * either the per-axis cap ({@link MAX_CANVAS_DIMENSION}) or the
 * area cap ({@link MAX_CANVAS_AREA}) is exceeded, both axes are multiplied by the
 * single largest factor `s ≤ 1` that brings the size within BOTH caps, so the
 * result keeps the requested aspect ratio.
 *
 * Returns the clamped integer dimensions, the applied `scale` (1 when nothing
 * was clamped), and a `clamped` flag. Never throws; non-finite or non-positive
 * inputs collapse to a 1×1 canvas.
 */
export function clampCanvasSize(width: number, height: number): ClampedCanvasSize {
  // Normalize: non-finite / non-positive → 1. Round to integer pixels, then
  // floor to ≥ 1 so a sub-0.5px request (which `Math.round` sends to 0) still
  // becomes a real 1×1 backing store instead of a blank 0×0 one — the unclamped
  // return path below trusts these to already be valid dimensions.
  const reqW = Number.isFinite(width) && width > 0 ? Math.max(1, Math.round(width)) : 1;
  const reqH = Number.isFinite(height) && height > 0 ? Math.max(1, Math.round(height)) : 1;

  // Per-axis factor: shrink until each axis is within the dimension cap.
  const dimScale = Math.min(
    1,
    MAX_CANVAS_DIMENSION / reqW,
    MAX_CANVAS_DIMENSION / reqH,
  );
  // Area factor: the area scales with s², so the linear factor is sqrt(cap/area).
  const area = reqW * reqH;
  const areaScale = area > MAX_CANVAS_AREA ? Math.sqrt(MAX_CANVAS_AREA / area) : 1;

  const scale = Math.min(dimScale, areaScale);
  if (scale >= 1) {
    return { width: reqW, height: reqH, scale: 1, clamped: false };
  }

  // Apply the uniform factor; floor to stay strictly within the caps, but never
  // below 1px on either axis.
  const w = Math.max(1, Math.floor(reqW * scale));
  const h = Math.max(1, Math.floor(reqH * scale));
  return { width: w, height: h, scale, clamped: true };
}
