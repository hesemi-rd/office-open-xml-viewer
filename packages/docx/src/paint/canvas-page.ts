import type {
  DocumentLayout,
  PaintNode,
} from '../layout/types.js';
import { orderedPagePaintNodes } from '../layout/page-graph.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import type { CanvasPaintContext, PaintPageOptions } from './types.js';

function paintNode(node: PaintNode, context: CanvasPaintContext): void {
  if (node.kind === 'drawing') paintDrawingLayout(node, context);
}

export async function paintLayoutPage(
  layout: DocumentLayout,
  pageIndex: number,
  target: HTMLCanvasElement | OffscreenCanvas,
  options: PaintPageOptions,
): Promise<void> {
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page ${pageIndex} is outside the layout`);
  const nodes = orderedPagePaintNodes(page);
  const ctx = target.getContext('2d') as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null;
  if (!ctx) throw new Error('Canvas 2D context is unavailable');

  const pixelScale = options.scale * options.dpr;
  target.width = Math.ceil(page.geometry.widthPt * pixelScale);
  target.height = Math.ceil(page.geometry.heightPt * pixelScale);
  ctx.save();
  try {
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, target.width, target.height);
    ctx.setTransform(pixelScale, 0, 0, pixelScale, 0, 0);
    for (const node of nodes) paintNode(node, { ctx, scale: options.scale, dpr: options.dpr });
  } finally {
    ctx.restore();
  }
}
