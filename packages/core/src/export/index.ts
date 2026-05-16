/**
 * Shared multi-page export primitives. Each format package re-exports a
 * format-specific wrapper that fills in the renderer; the heavy lifting
 * (offscreen canvas + PNG encoding) lives here so the three viewers stay
 * in sync.
 */

export interface PageBitmap {
  /** 0-based index in the source document. */
  index: number;
  /** PNG bytes for the rendered page / slide. */
  blob: Blob;
  /** Pixel dimensions of the PNG (multiplied by dpr already). */
  pixelWidth: number;
  pixelHeight: number;
  /** Page size in points (1/72 inch). Surfaced so callers writing their own
   *  PDF / SVG assembler have a single source of truth for page geometry. */
  pointWidth: number;
  pointHeight: number;
}

export interface RenderPageToCanvasContext {
  /** Total number of pages / slides / sheets to render. */
  pageCount: number;
  /** Draw a page onto a caller-supplied canvas. Implementations should
   *  size the canvas appropriately for the requested width / dpr. */
  renderPage: (canvas: HTMLCanvasElement, pageIndex: number, opts: { width: number; dpr: number }) => Promise<void>;
  /** Logical page size in points (1/72 inch). */
  pageSizeInPoints: (pageIndex: number) => { widthPt: number; heightPt: number };
}

export interface ExportPngOptions {
  /** Output width in CSS pixels (height derived from the page aspect ratio). */
  width?: number;
  /** Device pixel ratio. Default 2. */
  dpr?: number;
}

/** Render a single page to a PNG `Blob`. */
export async function renderPageToPng(
  ctx: RenderPageToCanvasContext,
  pageIndex: number,
  opts: ExportPngOptions = {},
): Promise<PageBitmap> {
  const width = opts.width ?? 1280;
  const dpr = opts.dpr ?? 2;
  const canvas = createOffscreen();
  await ctx.renderPage(canvas, pageIndex, { width, dpr });
  const blob = await canvasToPngBlob(canvas);
  const { widthPt, heightPt } = ctx.pageSizeInPoints(pageIndex);
  return {
    index: pageIndex,
    blob,
    pixelWidth: canvas.width,
    pixelHeight: canvas.height,
    pointWidth: widthPt,
    pointHeight: heightPt,
  };
}

/** Render every page as PNG. */
export async function renderAllPagesToPng(
  ctx: RenderPageToCanvasContext,
  opts: ExportPngOptions = {},
): Promise<PageBitmap[]> {
  const out: PageBitmap[] = [];
  for (let i = 0; i < ctx.pageCount; i++) {
    out.push(await renderPageToPng(ctx, i, opts));
  }
  return out;
}

function createOffscreen(): HTMLCanvasElement {
  // Always pick an HTMLCanvasElement (not OffscreenCanvas) so the renderers
  // can `instanceof`-check it and apply CSS sizing logic uniformly.
  const canvas = document.createElement('canvas');
  return canvas;
}

function canvasToPngBlob(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('canvas.toBlob returned null — encoder failed'));
        return;
      }
      resolve(blob);
    }, 'image/png');
  });
}
