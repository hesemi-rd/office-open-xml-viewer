import type {
  BlockLayoutAlgorithms,
  FlowLayout,
  FlowLayoutInput,
  LayoutRect,
  LayoutServices,
  ParagraphLayout,
  TableLayout,
} from './types.js';

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
  const blocks: Array<ParagraphLayout | TableLayout> = input.blocks.map((block) => (
    block.kind === 'paragraph'
      ? algorithms.layoutParagraph(block, services)
      : algorithms.layoutTable(block, services)
  ));

  return {
    source: input.source,
    container: input.container,
    blocks,
    flowBounds: unionBounds(blocks.map((block) => block.flowBounds), input.container.bounds),
    inkBounds: unionBounds(blocks.map((block) => block.inkBounds), input.container.bounds),
    clipBounds: input.container.bounds,
    advancePt: blocks.reduce((sum, block) => sum + block.advancePt, 0),
    ordinaryFlow: true,
  };
}
