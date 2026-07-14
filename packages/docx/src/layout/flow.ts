import type {
  BlockLayoutAlgorithms,
  FlowLayout,
  FlowLayoutInput,
  LayoutRect,
  LayoutServices,
  ParagraphLayout,
  TableLayout,
} from './types.js';
import { LayoutInvariantError } from './diagnostics.js';

function unionBounds(bounds: readonly LayoutRect[], fallback: LayoutRect): LayoutRect {
  if (bounds.length === 0) {
    return { xPt: fallback.xPt, yPt: fallback.yPt, widthPt: 0, heightPt: 0 };
  }
  const left = Math.min(...bounds.map((rect) => rect.xPt));
  const top = Math.min(...bounds.map((rect) => rect.yPt));
  const right = Math.max(...bounds.map((rect) => rect.xPt + rect.widthPt));
  const bottom = Math.max(...bounds.map((rect) => rect.yPt + rect.heightPt));
  return { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
}

export function layoutFlowBlocks(
  input: FlowLayoutInput,
  services: LayoutServices,
  algorithms: BlockLayoutAlgorithms,
): FlowLayout {
  const blocks: Array<ParagraphLayout | TableLayout> = [];
  let cursor = input.cursor;
  const containerBottom = input.container.bounds.yPt + input.container.bounds.heightPt;
  const containerRight = input.container.bounds.xPt + input.container.bounds.widthPt;

  for (const block of input.blocks) {
    const placement = {
      container: input.container,
      cursor,
      availableBounds: {
        xPt: input.container.bounds.xPt,
        yPt: cursor.yPt,
        widthPt: input.container.bounds.widthPt,
        heightPt: Math.max(0, containerBottom - cursor.yPt),
      },
    };
    const result = block.kind === 'paragraph'
      ? algorithms.layoutParagraph(block, placement, services)
      : algorithms.layoutTable(block, placement, services);
    if (result.layout.flowDomainId !== input.container.id) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `${result.layout.id} belongs to ${result.layout.flowDomainId}, not ${input.container.id}`,
      );
    }
    if (!Number.isFinite(result.nextCursor.xPt)
      || !Number.isFinite(result.nextCursor.yPt)
      || result.nextCursor.xPt < input.container.bounds.xPt
      || result.nextCursor.xPt > containerRight
      || result.nextCursor.yPt < cursor.yPt
      || result.nextCursor.yPt > containerBottom) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${result.layout.id} returned an invalid flow cursor`);
    }
    blocks.push(result.layout);
    cursor = result.nextCursor;
  }

  return {
    source: input.source,
    container: input.container,
    blocks,
    nextCursor: cursor,
    flowDomainId: input.container.id,
    flowBounds: unionBounds(blocks.map((block) => block.flowBounds), input.container.bounds),
    inkBounds: unionBounds(blocks.map((block) => block.inkBounds), input.container.bounds),
    clipBounds: input.container.bounds,
    advancePt: cursor.yPt - input.cursor.yPt,
    ordinaryFlow: true,
  };
}
