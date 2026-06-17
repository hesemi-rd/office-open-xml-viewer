import { defaultDpr, isHTMLCanvas, getCachedSvgImageByPath, type MathRenderer } from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions } from './types.js';
import { renderViewport, prepareWorksheetMath, worksheetHasUncachedMath } from './renderer.js';

/** What `prefetchImages` needs to decode one picture: the raster `imagePath`
 *  (also the cache key), its `mimeType`, and the optional svgBlip vector path. */
interface ImageRef {
  imagePath: string;
  mimeType: string;
  svgImagePath?: string;
}

/** Fetch one image's bytes by zip path and decode them to a drawable
 *  `CanvasImageSource`, preferring the Microsoft svgBlip vector original
 *  (MS-ODRAWXML). Unified across the top-level twoCellAnchor picture
 *  (`ImageAnchor`) and the `<xdr:grpSp>` leaf (`ShapeGeom` image) — both carry a
 *  raster `imagePath` fallback plus an optional `svgImagePath`. xlsx images have
 *  no `a:srcRect` crop, so the vector branch always applies when an svgBlip is
 *  present (cf. the contract's `!srcRect` gate).
 *
 *  Raster decodes to an `ImageBitmap` via `createImageBitmap`; the SVG vector
 *  original decodes to an `HTMLImageElement` via core's path-keyed
 *  `getCachedSvgImageByPath`, because `createImageBitmap` cannot rasterize SVG
 *  in every browser. Bytes are fetched lazily by zip path through `fetchImage`
 *  (twin of pptx/docx's `fetchImage`) instead of being inlined as base64. */
export async function decodeImageSource(
  imagePath: string,
  mimeType: string,
  svgImagePath: string | undefined,
  fetchImage: (path: string, mime: string) => Promise<Blob>,
): Promise<CanvasImageSource> {
  const decodeRaster = async (path: string, mime: string): Promise<CanvasImageSource> =>
    createImageBitmap(await fetchImage(path, mime));
  const dataIsSvg = mimeType === 'image/svg+xml';
  if (svgImagePath != null) {
    // Prefer the vector original; fall back to the raster fallback on decode
    // failure (or, when `imagePath` is itself the SVG, the SVG decoder again).
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
  imageCache: Map<string, CanvasImageSource>,
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
          });
        }
      }
    }
  }
  if (uncached.size === 0) return;
  await Promise.all(
    [...uncached.values()].map(async (ref) => {
      try {
        imageCache.set(
          ref.imagePath,
          await decodeImageSource(ref.imagePath, ref.mimeType, ref.svgImagePath, fetch),
        );
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
  imageCache: Map<string, CanvasImageSource>;
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

  target.width = Math.round(width * dpr);
  target.height = Math.round(height * dpr);
  // Set CSS display size so the browser renders at 1:1 device pixels (no browser-level scaling).
  // Without this, canvas.width=2400 on a DPR=2 display causes the canvas to be laid out at
  // 2400 CSS px, making all content appear blurry when viewed in a 1200 CSS px container.
  if (isHTMLCanvas(target)) {
    target.style.width = `${width}px`;
    target.style.height = `${height}px`;
  }

  const ctx = (target as HTMLCanvasElement).getContext('2d') as CanvasRenderingContext2D;
  ctx.scale(dpr, dpr);

  renderViewport(ctx, ws, styles, viewport, { ...opts, dpr, loadedImages: imageCache });
}
