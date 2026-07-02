import {
  defaultDpr,
  isHTMLCanvas,
  getCachedSvgImageByPath,
  decodeRasterOrMetafile,
  metafileRasterSize,
  EMU_PER_PT,
  type MathRenderer,
  type SrcRect,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions } from './types.js';
import { renderViewport, prepareWorksheetMath, worksheetHasUncachedMath } from './renderer.js';

/** What `prefetchImages` needs to decode one picture: the raster `imagePath`
 *  (also the cache key), its `mimeType`, the optional svgBlip vector path, and
 *  the picture's intended draw size in points (sizes a metafile raster; 0 ⇒
 *  decoder fallback). */
interface ImageRef {
  imagePath: string;
  mimeType: string;
  svgImagePath?: string;
  widthPt?: number;
  heightPt?: number;
  /** The picture's `<a:srcRect>` crop (§20.1.8.55), when present. Forces the
   *  raster decode (the crop math needs native bitmap pixels) and, for a
   *  metafile, scales the raster up to the full picture frame so the fractional
   *  crop lands correctly (see `metafileRasterSize`). */
  srcRect?: SrcRect | null;
}

/** Fetch one image's bytes by zip path and decode them to a drawable
 *  `CanvasImageSource`, preferring the Microsoft svgBlip vector original
 *  (MS-ODRAWXML). Unified across the top-level twoCellAnchor picture
 *  (`ImageAnchor`) and the `<xdr:grpSp>` leaf (`ShapeGeom` image) — both carry a
 *  raster `imagePath` fallback plus an optional `svgImagePath`. The svgBlip
 *  vector branch applies only when the picture is NOT cropped: with an
 *  `<a:srcRect>` crop (`hasCrop`) we force the raster, because the renderer's
 *  crop math needs the decoded bitmap's native pixel grid (an SVG element has
 *  none). Mirrors the pptx renderer's `!srcRect` vector gate.
 *
 *  Raster decodes to an `ImageBitmap` through core's
 *  {@link decodeRasterOrMetafile} (which content-sniffs the bytes: a WMF, which
 *  `createImageBitmap` can't decode, is rasterized by the shared minimal player
 *  at a size derived from `widthPt`/`heightPt`; a true EMF — or a WMF with no
 *  geometry — resolves to `null`, so the picture is skipped rather than
 *  crashing). The SVG vector original decodes to an `HTMLImageElement` via
 *  core's path-keyed `getCachedSvgImageByPath`, because `createImageBitmap`
 *  cannot rasterize SVG in every browser. Bytes are fetched lazily by zip path
 *  through `fetchImage` (twin of pptx/docx's `fetchImage`) instead of being
 *  inlined as base64.
 *
 *  Returns `null` for an unsupported metafile so the caller leaves the path
 *  uncached and the renderer skips a missing source. */
export async function decodeImageSource(
  imagePath: string,
  mimeType: string,
  svgImagePath: string | undefined,
  fetchImage: (path: string, mime: string) => Promise<Blob>,
  widthPt = 0,
  heightPt = 0,
  srcRect: SrcRect | null = null,
): Promise<CanvasImageSource | null> {
  const hasCrop = srcRect != null;
  const decodeRaster = async (path: string, mime: string): Promise<CanvasImageSource | null> => {
    // A cropped metafile must rasterize at its FULL picture frame, not the
    // visible sub-rect, so the fractional crop lands correctly; raster blips and
    // uncropped metafiles pass the box through unchanged.
    const sized = metafileRasterSize(mime, srcRect, widthPt, heightPt);
    return decodeRasterOrMetafile(await fetchImage(path, mime), {
      widthPt: sized.widthPt,
      heightPt: sized.heightPt,
    });
  };
  const dataIsSvg = mimeType === 'image/svg+xml';
  if (svgImagePath != null && !hasCrop) {
    // No crop: prefer the vector original; fall back to the raster on decode
    // failure (or, when `imagePath` is itself the SVG, the SVG decoder again).
    // A cropped picture skips this branch so the crop math (below, in the
    // renderer) runs on the raster bitmap's native pixel dimensions.
    try {
      return await getCachedSvgImageByPath(svgImagePath, fetchImage);
    } catch {
      return dataIsSvg
        ? getCachedSvgImageByPath(imagePath, fetchImage)
        : decodeRaster(imagePath, mimeType);
    }
  }
  if (dataIsSvg) {
    // svg-only picture with no separate `svgImagePath` field (defensive): the
    // raster decoder (createImageBitmap) can't rasterize SVG.
    return getCachedSvgImageByPath(imagePath, fetchImage);
  }
  return decodeRaster(imagePath, mimeType);
}

/** Collect every embedded image referenced by a worksheet and decode the ones
 *  not already in `imageCache`, storing each decoded `CanvasImageSource` under
 *  its zip `imagePath` (the renderer's lookup key). Images appear either as a
 *  top-level twoCellAnchor `<xdr:pic>` (in `ws.images`) or as a leaf inside an
 *  `<xdr:grpSp>` (a `ShapeGeom` with `type: 'image'`); BOTH are collected so the
 *  renderer never hits a missing source during the synchronous draw. De-duped
 *  by `imagePath` so a path shared across anchors is fetched + decoded once.
 *  A no-op when `fetchImage` is absent (no byte source). Per-image failures are
 *  swallowed so one broken picture doesn't sink the grid. */
export async function prefetchImages(
  ws: Worksheet,
  imageCache: Map<string, CanvasImageSource | null>,
  fetchImage: ((path: string, mime: string) => Promise<Blob>) | undefined,
): Promise<void> {
  if (!fetchImage) return;
  const fetch = fetchImage;
  const uncached = new Map<string, ImageRef>();
  if (ws.images) {
    for (const img of ws.images) {
      if (!imageCache.has(img.imagePath)) {
        uncached.set(img.imagePath, {
          imagePath: img.imagePath,
          mimeType: img.mimeType,
          svgImagePath: img.svgImagePath,
          // Saved EMU extent → pt sizes a metafile raster (0 ⇒ decoder fallback).
          widthPt: img.nativeExtCx > 0 ? img.nativeExtCx / EMU_PER_PT : 0,
          heightPt: img.nativeExtCy > 0 ? img.nativeExtCy / EMU_PER_PT : 0,
          // An `<a:srcRect>` crop forces the raster decode (native pixel grid)
          // and, for a metafile, the full-frame raster size.
          srcRect: img.srcRect ?? null,
        });
      }
    }
  }
  if (ws.shapeGroups) {
    for (const grp of ws.shapeGroups) {
      for (const shape of grp.shapes) {
        if (shape.geom.type === 'image' && !imageCache.has(shape.geom.imagePath)) {
          uncached.set(shape.geom.imagePath, {
            imagePath: shape.geom.imagePath,
            mimeType: shape.geom.mimeType,
            svgImagePath: shape.geom.svgImagePath,
            // Group's saved EMU extent scaled by the leaf's normalized w/h → pt.
            widthPt: grp.nativeExtCx > 0 ? (grp.nativeExtCx * shape.w) / EMU_PER_PT : 0,
            heightPt: grp.nativeExtCy > 0 ? (grp.nativeExtCy * shape.h) / EMU_PER_PT : 0,
            // A crop forces the raster decode (native pixel grid for the crop)
            // and, for a metafile, the full-frame raster size.
            srcRect: shape.geom.srcRect ?? null,
          });
        }
      }
    }
  }
  if (uncached.size === 0) return;
  await Promise.all(
    [...uncached.values()].map(async (ref) => {
      try {
        const src = await decodeImageSource(
          ref.imagePath,
          ref.mimeType,
          ref.svgImagePath,
          fetch,
          ref.widthPt,
          ref.heightPt,
          ref.srcRect,
        );
        // Cache the decode result keyed by path — INCLUDING a null for an
        // unsupported metafile (true EMF / geometry-less WMF). Storing the null
        // (matching pptx's getCachedBitmap) makes `imageCache.has(path)` short-
        // circuit the per-render prefetch, so the blip is sniffed ONCE instead of
        // re-fetched + re-sniffed every viewport frame. The renderer already
        // skips a falsy source, so a cached null draws nothing.
        imageCache.set(ref.imagePath, src);
      } catch {
        /* leave uncached; renderer skips a missing source */
      }
    }),
  );
}

export interface RenderDeps {
  ws: Worksheet;
  styles: ParsedWorkbook['styles'];
  /** Shared decoded-image cache, owned by the caller (workbook or worker). */
  imageCache: Map<string, CanvasImageSource | null>;
  math?: MathRenderer;
}

/** The full per-frame orchestration: preload uncached images, pre-rasterize
 *  equations, size the target, draw. Shared verbatim by the main-thread
 *  XlsxWorkbook and the render worker. */
export async function renderWorksheetViewport(
  deps: RenderDeps,
  target: HTMLCanvasElement | OffscreenCanvas,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): Promise<void> {
  const { ws, styles, imageCache } = deps;

  // ── Step 1: Preload any uncached image sources BEFORE touching the canvas.
  //
  // Images can appear either as top-level twoCellAnchor `<xdr:pic>` (captured
  // in `ws.images`) or as a leaf inside an `<xdr:grpSp>` (captured as a
  // ShapeGeom with `type: 'image'`); `prefetchImages` collects both, keyed by
  // zip `imagePath`, fetching bytes lazily via `opts.fetchImage`.
  //
  // Doing this *before* the canvas resize is critical for scroll smoothness:
  // setting `canvas.width` wipes the canvas, and an `await` after that wipe
  // yields to the browser's paint cycle, causing a visible white flash on
  // every scroll frame. By awaiting first (and only when there's something
  // uncached), the whole resize+draw runs synchronously in a single tick and
  // the old frame stays visible until the new one is ready.
  await prefetchImages(ws, imageCache, opts.fetchImage);

  // ── Step 1b: Pre-rasterize equations in shapes BEFORE the canvas resize,
  // for the same no-white-flash reason as the image preload. Gated on
  // `worksheetHasUncachedMath` so steady-state scroll/zoom frames take NO
  // await and stay fully synchronous — only the first frame that reveals new
  // equations pays the (idempotently cached) MathJax cost. Opt-in: skipped
  // entirely unless the caller supplies a `math` engine.
  if (deps.math && worksheetHasUncachedMath(ws)) {
    await prepareWorksheetMath(ws, deps.math);
  }

  // ── Step 2: Resize + draw, all synchronous from here.
  const dpr = opts.dpr ?? defaultDpr();
  const rawW = isHTMLCanvas(target) ? (target.clientWidth || 800) : target.width;
  const rawH = isHTMLCanvas(target) ? (target.clientHeight || 600) : target.height;
  const width = opts.width ?? rawW;
  const height = opts.height ?? rawH;

  // Resize only when the backing store dimensions actually change. Assigning
  // canvas.width/height re-allocates (and clears) the GPU backing store, so on a
  // steady-state scroll/zoom stream — where width/height/dpr are unchanged frame
  // to frame — re-assigning the same value wastes an allocation every frame
  // (improvement plan C4). The inner renderViewport starts with an explicit
  // clearRect + white fill, so nothing depends on the width-assignment's implicit
  // clear; skipping the same-size resize is safe.
  const bw = Math.round(width * dpr);
  const bh = Math.round(height * dpr);
  if (target.width !== bw) target.width = bw;
  if (target.height !== bh) target.height = bh;
  // Set CSS display size so the browser renders at 1:1 device pixels (no browser-level scaling).
  // Without this, canvas.width=2400 on a DPR=2 display causes the canvas to be laid out at
  // 2400 CSS px, making all content appear blurry when viewed in a 1200 CSS px container.
  if (isHTMLCanvas(target)) {
    const cssW = `${width}px`;
    const cssH = `${height}px`;
    if (target.style.width !== cssW) target.style.width = cssW;
    if (target.style.height !== cssH) target.style.height = cssH;
  }

  const ctx = (target as HTMLCanvasElement).getContext('2d') as CanvasRenderingContext2D;
  // Set the DPR transform absolutely rather than ctx.scale(dpr, dpr): when the
  // resize above is skipped the backing store is NOT re-created, so its transform
  // is not reset to identity, and a relative scale() would compound the dpr every
  // frame (progressive zoom). setTransform is idempotent whether or not the store
  // was reallocated.
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  renderViewport(ctx, ws, styles, viewport, { ...opts, dpr, loadedImages: imageCache });
}
