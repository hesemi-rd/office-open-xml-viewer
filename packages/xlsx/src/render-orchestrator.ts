import { defaultDpr, isHTMLCanvas, type MathRenderer } from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions } from './types.js';
import { renderViewport, prepareWorksheetMath, worksheetHasUncachedMath } from './renderer.js';

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
  const uncached: string[] = [];
  if (ws.images) {
    for (const img of ws.images) {
      if (!imageCache.has(img.dataUrl)) uncached.push(img.dataUrl);
    }
  }
  if (ws.shapeGroups) {
    for (const grp of ws.shapeGroups) {
      for (const shape of grp.shapes) {
        if (shape.geom.type === 'image' && !imageCache.has(shape.geom.dataUrl)) {
          uncached.push(shape.geom.dataUrl);
        }
      }
    }
  }
  if (uncached.length > 0) {
    await Promise.all(
      uncached.map(async (url) => {
        const blob = await (await fetch(url)).blob();
        const bmp = await createImageBitmap(blob);
        imageCache.set(url, bmp);
      }),
    ).catch(() => { /* swallow image failures so the grid still renders */ });
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
