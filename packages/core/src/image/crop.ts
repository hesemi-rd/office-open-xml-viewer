// Shared `<a:srcRect>` source-rectangle crop (ECMA-376 §20.1.8.55) for the docx,
// pptx and xlsx renderers. The crop is a fraction of the image's NATIVE pixel
// grid, so the only per-renderer concern is decoding the image at full source
// size first — for a metafile that means rasterizing the whole picture FRAME
// (see `metafileRasterSize`), since the player maps the EMF/WMF frame onto the
// raster and the crop is relative to that frame. Centralised here so all three
// renderers crop identically (previously triplicated, and metafiles diverged).

import { isMetafileMime } from './wmf';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** A `<a:srcRect>` crop: fractional insets (0..1) measured inward from each edge.
 *  The visible region is `[l, t, 1−r, 1−b]` of the source. (Mirrors the Rust
 *  `SrcRect` the parsers already emit, divided to 0..1 fractions.) */
export interface SrcRect {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** Native pixel size of a decoded image (ImageBitmap exposes `width`/`height`;
 *  an `<img>` element exposes `naturalWidth`/`naturalHeight`). */
export function imageNaturalSize(img: CanvasImageSource): { w: number; h: number } {
  const el = img as {
    naturalWidth?: number;
    naturalHeight?: number;
    width?: number;
    height?: number;
  };
  const w = el.naturalWidth || (typeof el.width === 'number' ? el.width : 0) || 0;
  const h = el.naturalHeight || (typeof el.height === 'number' ? el.height : 0) || 0;
  return { w, h };
}

/** The 9-arg `drawImage` source rectangle for an `<a:srcRect>` crop, or `null`
 *  when there is no (non-empty) crop or the image reports no native size.
 *
 *  `sx = l·W`, `sy = t·H`, `sw = (1−l−r)·W`, `sh = (1−t−b)·H`, each clamped: a
 *  negative (overscan) edge degrades to 0 and the slice is kept ≥ 1px. Callers
 *  that need the rect for auxiliary paints (e.g. pptx effect passes) call this
 *  directly; the common path uses {@link drawImageCropped}. */
export function cropSourceRect(
  img: CanvasImageSource,
  srcRect: SrcRect | null | undefined,
): { sx: number; sy: number; sw: number; sh: number } | null {
  if (!srcRect || !(srcRect.l || srcRect.t || srcRect.r || srcRect.b)) return null;
  const { w, h } = imageNaturalSize(img);
  if (w <= 0 || h <= 0) return null;
  const c01 = (v: number): number => Math.max(0, Math.min(1, v));
  const sx = c01(srcRect.l) * w;
  const sy = c01(srcRect.t) * h;
  return {
    sx,
    sy,
    sw: Math.max(1, w - sx - c01(srcRect.r) * w),
    sh: Math.max(1, h - sy - c01(srcRect.b) * h),
  };
}

/** Draw `img` into the destination box `[dx, dy, dw, dh]`, honoring an optional
 *  `<a:srcRect>` crop. The destination box is unchanged — the visible slice is
 *  stretched to fill it (the 9-arg `drawImage` behavior). Crop applies to raster
 *  blips AND metafiles alike: a cropped metafile must have been rasterized at its
 *  full frame via {@link metafileRasterSize}, so its bitmap is the full source. */
export function drawImageCropped(
  ctx: AnyCtx,
  img: CanvasImageSource,
  srcRect: SrcRect | null | undefined,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
): void {
  const c = cropSourceRect(img, srcRect);
  if (c) ctx.drawImage(img, c.sx, c.sy, c.sw, c.sh, dx, dy, dw, dh);
  else ctx.drawImage(img, dx, dy, dw, dh);
}

/** Raster target size (pt) for decoding an embedded image. A raster blip decodes
 *  at its native pixel grid, so its display box passes through. A metafile
 *  (WMF/EMF) with an `<a:srcRect>` crop must be rasterized at its FULL picture
 *  frame, not the visible sub-rectangle — the player maps the frame onto the
 *  raster (see `playEmf`), and the crop is relative to that frame. Scale the box
 *  up by `1/(1−l−r)` and `1/(1−t−b)` so the rasterised frame and the fractional
 *  crop align (e.g. one composite EMF cropped into subfigures). Uncropped
 *  metafiles and all raster blips pass the box through unchanged. */
export function metafileRasterSize(
  mimeType: string,
  srcRect: SrcRect | null | undefined,
  widthPt: number,
  heightPt: number,
): { widthPt: number; heightPt: number } {
  if (!srcRect || !isMetafileMime(mimeType)) return { widthPt, heightPt };
  const fracW = Math.max(0.01, 1 - srcRect.l - srcRect.r);
  const fracH = Math.max(0.01, 1 - srcRect.t - srcRect.b);
  return { widthPt: widthPt / fracW, heightPt: heightPt / fracH };
}
