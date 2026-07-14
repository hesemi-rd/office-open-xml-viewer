import type {
  DocumentLayout,
  PageLayerId,
  PaintNode,
} from '../layout/types.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import type { CanvasPaintContext, PaintPageOptions } from './types.js';

function paintNode(node: PaintNode, context: CanvasPaintContext): void {
  if (node.kind === 'drawing') paintDrawingLayout(node, context);
}

const PAGE_LAYER_IDS = [
  'background',
  'behindText',
  'header',
  'body',
  'notes',
  'front',
  'footer',
] as const satisfies readonly PageLayerId[];

function orderedPaintNodes(page: DocumentLayout['pages'][number]): readonly PaintNode[] {
  const nodes = new Map<string, { layer: PageLayerId; node: PaintNode }>();
  for (const layer of PAGE_LAYER_IDS) {
    for (const node of page.layers[layer]) {
      if (nodes.has(node.id)) throw new Error(`Duplicate paint node ${node.id}`);
      nodes.set(node.id, { layer, node });
    }
  }

  const painted = new Set<string>();
  const ordered: PaintNode[] = [];
  for (const entry of page.layers.paintOrder) {
    const target = nodes.get(entry.nodeId);
    if (!target) throw new Error(`Missing paint node ${entry.nodeId}`);
    if (target.layer !== entry.layer) {
      throw new Error(`Paint node ${entry.nodeId} belongs to ${target.layer}, not ${entry.layer}`);
    }
    if (painted.has(entry.nodeId)) throw new Error(`Duplicate paint reference ${entry.nodeId}`);
    painted.add(entry.nodeId);
    ordered.push(target.node);
  }
  if (painted.size !== nodes.size) {
    const missing = [...nodes.keys()].find((id) => !painted.has(id));
    throw new Error(`Missing paint-order reference for ${missing ?? '<unknown>'}`);
  }
  return ordered;
}

export async function paintLayoutPage(
  layout: DocumentLayout,
  pageIndex: number,
  target: HTMLCanvasElement | OffscreenCanvas,
  options: PaintPageOptions,
): Promise<void> {
  const page = layout.pages[pageIndex];
  if (!page) throw new RangeError(`Page ${pageIndex} is outside the layout`);
  const nodes = orderedPaintNodes(page);
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
