import type { CellElement, DocParagraph, DocTable, DocTableCell } from '../types.js';
import type { FlowFragment } from '../layout-fragments.js';
import type { ParagraphLayout } from './types.js';
import { paragraphGapPt } from './paragraph-spacing.js';
import { paragraphFragmentAdvancePt, tableFragmentHeightPt } from '../layout-fragments.js';
import {
  resolveParagraphBorderEdges,
  type ParagraphBorderEdges,
} from './paragraph-border-adjacency.js';

export type AcquireNestedCellBlocks<State> = (
  cell: DocTableCell,
  table: DocTable,
  cellTotalWidthPt: number,
  outerState: State,
  sourcePath: readonly number[],
) => readonly FlowFragment[];

export interface TableCellBlockAcquisitionDependencies<State> {
  resolveContentWidthPt(cell: DocTableCell, table: DocTable, totalWidthPt: number): number;
  createCellState(outerState: State, contentWidthPt: number, cell: DocTableCell): State;
  acquireParagraph(
    state: State,
    paragraph: DocParagraph,
    contentWidthPt: number,
    sourcePath: readonly number[],
    paragraphBorderEdges: ParagraphBorderEdges,
  ): ParagraphLayout;
  acquireNestedTable(
    state: State,
    table: DocTable,
    contentWidthPt: number,
    sourcePath: readonly number[],
    continuation: Readonly<{ fromPrevious: boolean; onNext: boolean }>,
    acquireNestedCellBlocks: AcquireNestedCellBlocks<State>,
  ): FlowFragment;
  advanceState(state: State, advancePt: number): void;
}

export interface AcquireTableCellBlocksInput<State> {
  readonly cell: DocTableCell;
  readonly table: DocTable;
  readonly cellTotalWidthPt: number;
  readonly outerState: State;
  readonly sourcePath: readonly number[];
}

export interface RetainedCellBlockPlacement {
  readonly blockPlacements: readonly Readonly<{ offsetPt: number; advancePt: number }>[];
  readonly contentTranslationPt: number;
  readonly inkBlock: Readonly<{ topPt: number; heightPt: number }>;
}

export function isStructuralTrailingParagraph(
  content: readonly CellElement[],
  index: number,
): boolean {
  if (index !== content.length - 1 || index === 0) return false;
  const current = content[index];
  const previous = content[index - 1];
  return current?.type === 'paragraph'
    && previous?.type === 'table'
    && current.runs.length === 0;
}

/**
 * Resolve the immutable point-space placement of a cell's retained block tree.
 * Paragraph edge spacing participates in top flow, but center/bottom align the
 * ink block between the resolved cell margins (ECMA-376 §17.4.83). Nested
 * tables use the same document-order fold rather than a table-specific branch.
 */
export function resolveRetainedCellBlockPlacement(
  cell: DocTableCell,
  table: DocTable,
  blocks: readonly FlowFragment[],
  boxHeightPt: number,
): RetainedCellBlockPlacement {
  const blockPlacements: Array<{ offsetPt: number; advancePt: number }> = [];
  let cursorPt = 0;
  let previousParagraph: DocParagraph | null = null;
  let previousAfterPt = 0;
  let firstInkTopPt: number | undefined;
  let lastInkBottomPt = 0;

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]!;
    const element = cell.content[index];
    const structural = isStructuralTrailingParagraph(cell.content, index);
    if (block.kind === 'paragraph' && element?.type === 'paragraph') {
      const paragraph: DocParagraph = element;
      const blockBeforePt = block.spacing?.beforePt ?? 0;
      const blockAfterPt = block.spacing?.afterPt ?? 0;
      const gapPt = previousParagraph
        ? paragraphGapPt(
            previousParagraph,
            paragraph,
            previousAfterPt,
            blockBeforePt,
          )
        : index === 0 || cell.content[index - 1]?.type === 'table'
          ? blockBeforePt
          : 0;
      const lineBlockPt = Math.max(
        0,
        paragraphFragmentAdvancePt(block) - blockBeforePt - blockAfterPt,
      );
      const offsetPt = cursorPt + gapPt;
      blockPlacements.push({ offsetPt, advancePt: lineBlockPt });
      cursorPt = offsetPt + lineBlockPt;
      if (!structural) {
        firstInkTopPt ??= offsetPt;
        lastInkBottomPt = cursorPt;
      }
      previousParagraph = paragraph;
      previousAfterPt = blockAfterPt;
      continue;
    }

    if (previousParagraph) cursorPt += previousAfterPt;
    const advancePt = block.kind === 'table'
      ? ('flowBounds' in block ? block.advancePt : tableFragmentHeightPt(block))
      : paragraphFragmentAdvancePt(block);
    blockPlacements.push({ offsetPt: cursorPt, advancePt });
    firstInkTopPt ??= cursorPt;
    cursorPt += advancePt;
    lastInkBottomPt = cursorPt;
    previousParagraph = null;
    previousAfterPt = 0;
  }

  const inkTopPt = firstInkTopPt ?? 0;
  const inkHeightPt = Math.max(0, lastInkBottomPt - inkTopPt);
  const marginTopPt = cell.marginTop ?? table.cellMarginTop ?? 0;
  const marginBottomPt = cell.marginBottom ?? table.cellMarginBottom ?? 0;
  const contentHeightPt = boxHeightPt - marginTopPt - marginBottomPt;
  const alignedInkTopPt = cell.vAlign === 'center'
    ? marginTopPt + (contentHeightPt - inkHeightPt) / 2
    : cell.vAlign === 'bottom'
      ? boxHeightPt - marginBottomPt - inkHeightPt
      : marginTopPt + inkTopPt;
  const contentTranslationPt = cell.vAlign === 'top'
    ? marginTopPt
    : alignedInkTopPt - inkTopPt;

  return {
    blockPlacements: Object.freeze(blockPlacements.map((placement) => Object.freeze(placement))),
    contentTranslationPt,
    inkBlock: Object.freeze({ topPt: inkTopPt, heightPt: inkHeightPt }),
  };
}

/**
 * Acquire a cell's recursive paragraph/table block tree. Table pagination owns
 * the injected geometry callbacks; this module owns document-order recursion and
 * the invariant that every paragraph becomes one self-contained ParagraphLayout.
 */
export function acquireTableCellBlocks<State>(
  input: AcquireTableCellBlocksInput<State>,
  dependencies: TableCellBlockAcquisitionDependencies<State>,
): readonly FlowFragment[] {
  const { cell, table, cellTotalWidthPt, outerState, sourcePath } = input;
  const contentWidthPt = dependencies.resolveContentWidthPt(cell, table, cellTotalWidthPt);
  const cellState = dependencies.createCellState(outerState, contentWidthPt, cell);
  const blocks: FlowFragment[] = [];

  for (let cellElementIndex = 0; cellElementIndex < cell.content.length; cellElementIndex += 1) {
    const element = cell.content[cellElementIndex];
    if (!element) continue;
    const elementPath = [...sourcePath, cellElementIndex];
    if (element.type === 'paragraph') {
      const previousElement = cell.content[cellElementIndex - 1];
      const nextElement = cell.content[cellElementIndex + 1];
      const paragraph: DocParagraph = element;
      const block = dependencies.acquireParagraph(
        cellState,
        paragraph,
        contentWidthPt,
        elementPath,
        resolveParagraphBorderEdges(
          previousElement?.type === 'paragraph'
            ? previousElement : null,
          paragraph,
          nextElement?.type === 'paragraph'
            ? nextElement : null,
        ),
      );
      blocks.push(block);
      dependencies.advanceState(cellState, block.advancePt);
      continue;
    }

    const inner: DocTable = element;
    const nestedSlice = element as typeof element & {
      nestedSliceContinuesFromPrevious?: boolean;
      nestedSliceContinuesOnNext?: boolean;
    };
    blocks.push(dependencies.acquireNestedTable(
      cellState,
      inner,
      contentWidthPt,
      elementPath,
      {
        fromPrevious: nestedSlice.nestedSliceContinuesFromPrevious ?? false,
        onNext: nestedSlice.nestedSliceContinuesOnNext ?? false,
      },
      (nestedCell, nestedTable, widthPt, nestedOuterState, nestedPath) =>
        acquireTableCellBlocks({
          cell: nestedCell,
          table: nestedTable,
          cellTotalWidthPt: widthPt,
          outerState: nestedOuterState,
          sourcePath: nestedPath,
        }, dependencies),
    ));
  }

  return blocks;
}
