import { defaultDpr, isHTMLCanvas, getCachedSvgImage, type MathRenderer } from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions } from './types.js';
import { renderViewport, prepareWorksheetMath, worksheetHasUncachedMath } from './renderer.js';

/** Decode one image element to a drawable `CanvasImageSource`, preferring the
 *  Microsoft svgBlip vector original (MS-ODRAWXML). Unified across the top-level
 *  twoCellAnchor picture (`ImageAnchor`) and the `<xdr:grpSp>` leaf
 *  (`ShapeGeom` image) — both carry a raster `dataUrl` fallback plus an optional
 *  `svgDataUrl`. xlsx images have no `a:srcRect` crop, so the vector branch
 *  always applies when an svgBlip is present (cf. the contract's `!srcRect`
 *  gate). SVG decodes through `getCachedSvgImage` (an `<img>`) because
 *  `createImageBitmap` cannot rasterize SVG in every browser. */
async function decodeImageSource(
  dataUrl: string,
  svgDataUrl?: string,
): Promise<CanvasImageSource> {
  const decodeRaster = async (url: string): Promise<CanvasImageSource> =>
    createImageBitmap(await (await fetch(url)).blob());
  const dataIsSvg = dataUrl.startsWith('data:image/svg+xml');
  if (svgDataUrl != null) {
    // Prefer the vector original; fall back to the raster fallback on decode
    // failure (or, when `dataUrl` is itself the SVG, the SVG decoder again).
    try {
      return await getCachedSvgImage(svgDataUrl);
    } catch {
      return dataIsSvg ? getCachedSvgImage(dataUrl) : decodeRaster(dataUrl);
    }
  }
  if (dataIsSvg) {
    // svg-only picture with no separate `svgDataUrl` field (defensive): the
    // raster decoder (createImageBitmap) can't rasterize SVG.
    return getCachedSvgImage(dataUrl);
  }
  return decodeRaster(dataUrl);
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

  // ── Step 1: Preload any uncached image bitmaps BEFORE touching the canvas.
  //
  // Images can appear either as top-level twoCellAnchor `<xdr:pic>` (captured
  // in `ws.images`) or as a leaf inside an `<xdr:grpSp>` (captured as a
  // ShapeGeom with `type: 'image'`). We collect both so the renderer never
  // hits a missing bitmap during the synchronous draw pass.
  //
  // Doing this *before* the canvas resize is critical for scroll smoothness:
  // setting `canvas.width` wipes the canvas, and an `await` after that wipe
  // yields to the browser's paint cycle, causing a visible white flash on
  // every scroll frame. By awaiting first (and only when there's something
  // uncached), the whole resize+draw runs synchronously in a single tick and
  // the old frame stays visible until the new one is ready.
  // The cache is keyed by `dataUrl` (the renderer's lookup key); the decoded
  // source may come from `svgDataUrl` (preferred) instead. De-dup by `dataUrl`.
  const uncached = new Map<string, string | undefined>();
  if (ws.images) {
    for (const img of ws.images) {
      if (!imageCache.has(img.dataUrl)) uncached.set(img.dataUrl, img.svgDataUrl);
    }
  }
  if (ws.shapeGroups) {
    for (const grp of ws.shapeGroups) {
      for (const shape of grp.shapes) {
        if (shape.geom.type === 'image' && !imageCache.has(shape.geom.dataUrl)) {
          uncached.set(shape.geom.dataUrl, shape.geom.svgDataUrl);
        }
      }
    }
  }
  if (uncached.size > 0) {
    await Promise.all(
      [...uncached].map(async ([dataUrl, svgDataUrl]) => {
        // Swallow per-image failures so one broken picture doesn't sink the grid.
        try {
          imageCache.set(dataUrl, await decodeImageSource(dataUrl, svgDataUrl));
        } catch { /* leave uncached; renderer skips a missing source */ }
      }),
    );
  }

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
