import type { RetainedTableAcquisition } from './table-acquisition.js';
import { sliceParagraphLayout } from './paragraph.js';
import { layoutTable, measureTableCellBlockFlowHeightPt } from './table.js';
import type {
  FlowBlockPlacement,
  FloatingTablePlacementLayout,
  LayoutServices,
  ParagraphLayout,
  TableCellBlockInput,
  TableCellLayout,
  TableCellLayoutInput,
  TableLayout,
  TableLayoutInput,
  TableRowLayout,
  TableRowLayoutInput,
} from './types.js';

export type TableFragmentOwnership = 'source' | 'repeated-header';

export type BlockContinuationRange =
  | Readonly<{
      kind: 'paragraph';
      blockIndex: number;
      lineStart: number;
      lineEnd: number;
    }>
  | Readonly<{
      kind: 'nested-table';
      blockIndex: number;
      childFragmentIndex: number;
    }>
  | Readonly<{
      kind: 'whole';
      blockIndex: number;
    }>;

export interface TableCellFragmentLayout extends TableCellLayout {
  readonly contentRanges: readonly BlockContinuationRange[];
  /** A page-local paint role; the source w:vMerge value remains unchanged. */
  readonly visualMergeOwnership?: 'continuation';
}

export interface TableRowFragmentLayout extends TableRowLayout {
  readonly logicalRowIndex: number;
  readonly fragmentIndex: number;
  readonly ownership: TableFragmentOwnership;
  readonly occurrenceId: string;
  readonly physicalPageIndex: number;
  readonly displayPageNumber: number;
  readonly cells: readonly TableCellFragmentLayout[];
}

export interface TableFragmentLayout extends TableLayout {
  readonly rows: readonly TableRowFragmentLayout[];
  readonly floatingTables: readonly FloatingTablePlacementLayout[];
}

interface TableCellFragmentCursor {
  readonly blockIndex: number;
  readonly paragraphLineStart: number;
  readonly nestedCursor: TableFragmentCursor | null;
  readonly nestedFragmentIndex: number;
}

export interface TableFragmentCursor {
  readonly rowIndex: number;
  readonly rowFragmentIndex: number;
  readonly cells: readonly TableCellFragmentCursor[];
}

export interface TableFragmentPageContext {
  readonly physicalPageIndex: number;
  readonly displayPageNumber: number;
  readonly occurrenceId: string;
}

export interface PageDependentTableBlockRequest {
  readonly logicalRowIndex: number;
  readonly logicalCellIndex: number;
  readonly sourceBlockIndex: number;
  readonly ownership: TableFragmentOwnership;
  readonly page: TableFragmentPageContext;
  readonly acquired: ParagraphLayout | TableLayout;
}

export interface TableFragmentContext {
  readonly availableHeightPt: number;
  readonly freshPageHeightPt: number;
  readonly placement: FlowBlockPlacement;
  readonly services: LayoutServices;
  readonly compatibility: 'word' | 'standard';
  /** Deterministic policy for floating rows taller than a fresh page band. */
  readonly oversizedRowPolicy?: 'split' | 'atomic';
  readonly page: TableFragmentPageContext;
  /** Reacquire only content whose destination page can change its geometry. */
  readonly reacquirePageDependentBlock?: (
    request: PageDependentTableBlockRequest,
  ) => ParagraphLayout | TableLayout;
}

export interface TableFragmentResult {
  readonly fragment: TableFragmentLayout | null;
  readonly nextCursor: TableFragmentCursor | null;
  readonly requiresFreshPage: boolean;
}

interface SelectedCell {
  readonly input: TableCellLayoutInput;
  readonly range: readonly BlockContinuationRange[];
  readonly next: TableCellFragmentCursor;
  readonly complete: boolean;
}

interface SelectedRow {
  readonly input: TableRowLayoutInput;
  readonly logicalRowIndex: number;
  readonly fragmentIndex: number;
  readonly ownership: TableFragmentOwnership;
  readonly ranges: readonly (readonly BlockContinuationRange[])[];
  readonly clipAtPageEnd?: boolean;
}

const EPSILON_PT = 0.0001;

function emptyCellCursor(): TableCellFragmentCursor {
  return Object.freeze({
    blockIndex: 0,
    paragraphLineStart: 0,
    nestedCursor: null,
    nestedFragmentIndex: 0,
  });
}

export function startTableFragmentCursor(): TableFragmentCursor {
  return Object.freeze({ rowIndex: 0, rowFragmentIndex: 0, cells: Object.freeze([]) });
}

function leadingHeaderCount(input: TableLayoutInput): number {
  let count = 0;
  while (input.rows[count]?.repeatedHeader === true) count += 1;
  return count;
}

function paginationRowHeight(source: RetainedTableAcquisition, rowIndex: number): number {
  const row = source.layout.rows[rowIndex];
  if (!row) return 0;
  // A vertically merged owner's content requirement and its physical track can
  // land on different source rows. Considering both permits a legal boundary
  // through the span without rewriting restart/continue semantics.
  return Math.max(0, row.heightPt, row.contentHeightPt);
}

function paginationRowHeightForOccurrence(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  rowIndex: number,
  context: TableFragmentContext,
): number {
  if (row === source.input.rows[rowIndex]) return paginationRowHeight(source, rowIndex);
  const occurrence = layoutTable({
    ...source.input,
    id: `${source.input.id}:row-occurrence:${context.page.occurrenceId}:${row.logicalRowIndex}`,
    rows: [row],
  }, context.placement, context.services).layout;
  return Math.max(0, occurrence.rows[0]?.heightPt ?? occurrence.advancePt);
}

function rowRanges(row: TableRowLayoutInput): readonly (readonly BlockContinuationRange[])[] {
  return row.cells.map((cell) => cell.blocks.map((block) => ({
    kind: 'whole' as const,
    blockIndex: block.sourceBlockIndex,
  })));
}

function rowForOccurrence(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  context: TableFragmentContext,
): TableRowLayoutInput {
  const reacquire = context.reacquirePageDependentBlock;
  if (!reacquire || !row.cells.some((cell) => (
    cell.blocks.some((block) => block.pageDependent === true)
  ))) return row;
  return {
    ...row,
    cells: row.cells.map((cell, logicalCellIndex) => ({
      ...cell,
      blocks: cell.blocks.map((block) => block.pageDependent === true
        ? {
            ...block,
            layout: reacquire({
              logicalRowIndex: row.logicalRowIndex,
              logicalCellIndex,
              sourceBlockIndex: block.sourceBlockIndex,
              ownership,
              page: context.page,
              acquired: block.layout,
            }),
          }
        : block),
    })),
  };
}

function selectedWholeRow(
  row: TableRowLayoutInput,
  ownership: TableFragmentOwnership,
  fragmentIndex = 0,
  clipAtPageEnd = false,
): SelectedRow {
  return {
    input: row,
    logicalRowIndex: row.logicalRowIndex,
    fragmentIndex,
    ownership,
    ranges: rowRanges(row),
    ...(clipAtPageEnd ? { clipAtPageEnd: true } : {}),
  };
}

function paragraphSlice(
  paragraph: ParagraphLayout,
  start: number,
  end: number,
): ParagraphLayout {
  return sliceParagraphLayout(paragraph, {
    lineStart: start,
    lineEnd: end,
    continuesFromPrevious: start > 0,
    continuesOnNext: end < paragraph.lines.length,
  });
}

function selectParagraph(
  paragraph: ParagraphLayout,
  sourceBlockIndex: number,
  start: number,
  selectedBlocks: readonly TableCellBlockInput[],
  availableHeightPt: number,
): Readonly<{
  block: TableCellBlockInput | null;
  range: BlockContinuationRange | null;
  lineEnd: number;
  advancePt: number;
}> {
  let selected: ParagraphLayout | null = null;
  let lineEnd = start;
  for (let candidateEnd = start + 1; candidateEnd <= paragraph.lines.length; candidateEnd += 1) {
    const candidate = paragraphSlice(paragraph, start, candidateEnd);
    const candidateBlock = { layout: candidate, sourceBlockIndex } as const;
    if (measureTableCellBlockFlowHeightPt([...selectedBlocks, candidateBlock])
      > availableHeightPt + EPSILON_PT) break;
    selected = candidate;
    lineEnd = candidateEnd;
  }
  if (!selected) return { block: null, range: null, lineEnd: start, advancePt: 0 };
  return {
    block: { layout: selected, sourceBlockIndex },
    range: { kind: 'paragraph', blockIndex: sourceBlockIndex, lineStart: start, lineEnd },
    lineEnd,
    advancePt: selected.advancePt,
  };
}

function selectCell(
  source: RetainedTableAcquisition,
  cell: TableCellLayoutInput,
  cursor: TableCellFragmentCursor,
  availableContentHeightPt: number,
  context: TableFragmentContext,
): SelectedCell {
  if (cell.verticalMerge === 'continue') {
    return { input: cell, range: [], next: cursor, complete: true };
  }
  const blocks: TableCellBlockInput[] = [];
  const range: BlockContinuationRange[] = [];
  let blockIndex = cursor.blockIndex;
  let paragraphLineStart = cursor.paragraphLineStart;
  let nestedCursor = cursor.nestedCursor;
  let nestedFragmentIndex = cursor.nestedFragmentIndex;

  while (blockIndex < cell.blocks.length) {
    const sourceBlock = cell.blocks[blockIndex]!;
    const child = sourceBlock.layout;
    if (child.kind === 'paragraph') {
      if (sourceBlock.structuralTrailing) {
        blocks.push(sourceBlock);
        range.push({ kind: 'whole', blockIndex: sourceBlock.sourceBlockIndex });
        blockIndex += 1;
        paragraphLineStart = 0;
        continue;
      }
      const selected = selectParagraph(
        child,
        sourceBlock.sourceBlockIndex,
        paragraphLineStart,
        blocks,
        availableContentHeightPt,
      );
      if (!selected.block || !selected.range) break;
      blocks.push({ ...selected.block, ...(sourceBlock.structuralTrailing
        ? { structuralTrailing: true }
        : {}) });
      range.push(selected.range);
      if (selected.lineEnd < child.lines.length) {
        paragraphLineStart = selected.lineEnd;
        break;
      }
      blockIndex += 1;
      paragraphLineStart = 0;
      continue;
    }

    const nested = source.nestedById[child.id];
    if (nested) {
      const remainingPt = Math.max(
        0,
        availableContentHeightPt - measureTableCellBlockFlowHeightPt(blocks),
      );
      const nestedResult = takeTableFragment(nested, nestedCursor ?? startTableFragmentCursor(), {
        ...context,
        availableHeightPt: remainingPt,
        placement: {
          ...context.placement,
          availableBounds: { ...context.placement.availableBounds, heightPt: remainingPt },
        },
      });
      if (!nestedResult.fragment) break;
      blocks.push({ layout: nestedResult.fragment, sourceBlockIndex: sourceBlock.sourceBlockIndex });
      range.push({
        kind: 'nested-table',
        blockIndex: sourceBlock.sourceBlockIndex,
        childFragmentIndex: nestedFragmentIndex,
      });
      if (nestedResult.nextCursor) {
        nestedCursor = nestedResult.nextCursor;
        nestedFragmentIndex += 1;
        break;
      }
      blockIndex += 1;
      nestedCursor = null;
      nestedFragmentIndex = 0;
      continue;
    }

    if (measureTableCellBlockFlowHeightPt([...blocks, sourceBlock])
      > availableContentHeightPt + EPSILON_PT) break;
    blocks.push(sourceBlock);
    range.push({ kind: 'whole', blockIndex: sourceBlock.sourceBlockIndex });
    blockIndex += 1;
  }

  const complete = blockIndex >= cell.blocks.length;
  return {
    input: { ...cell, blocks },
    range,
    next: Object.freeze({ blockIndex, paragraphLineStart, nestedCursor, nestedFragmentIndex }),
    complete,
  };
}

function partialRow(
  source: RetainedTableAcquisition,
  row: TableRowLayoutInput,
  cursor: TableFragmentCursor,
  availableHeightPt: number,
  context: TableFragmentContext,
): Readonly<{
  selected: SelectedRow | null;
  next: TableFragmentCursor;
  complete: boolean;
}> {
  const cellCursors = row.cells.map((_, index) => cursor.cells[index] ?? emptyCellCursor());
  const verticalInsetsPt = Math.max(0, ...row.cells.map((cell) => (
    cell.margins.topPt + cell.margins.bottomPt
  )));
  // A one-row fragment owns both outer cell-spacing bands. Reserve them before
  // selecting legal child boundaries so layoutTable cannot grow past the page.
  const spacingInsetsPt = Math.max(0, row.cellSpacingPt) * 2;
  const availableContentHeightPt = Math.max(
    0,
    availableHeightPt - verticalInsetsPt - spacingInsetsPt,
  );
  const selectedCells = row.cells.map((cell, index) => selectCell(
    source,
    cell,
    cellCursors[index]!,
    availableContentHeightPt,
    context,
  ));
  const madeProgress = selectedCells.some((cell, index) => (
    cell.next.blockIndex !== cellCursors[index]?.blockIndex
    || cell.next.paragraphLineStart !== cellCursors[index]?.paragraphLineStart
    || cell.next.nestedFragmentIndex !== cellCursors[index]?.nestedFragmentIndex
  ));
  if (!madeProgress) return { selected: null, next: cursor, complete: false };

  const complete = selectedCells.every((cell) => cell.complete);
  if (complete && cursor.rowFragmentIndex === 0) {
    return {
      selected: selectedWholeRow(row, 'source'),
      next: Object.freeze({
        rowIndex: cursor.rowIndex + 1,
        rowFragmentIndex: 0,
        cells: Object.freeze([]),
      }),
      complete: true,
    };
  }
  // Reaching this branch means content genuinely continues from or onto another
  // fragment: a fully retained first fragment returned as a whole row above.
  // Authored exact/atLeast height constrains that logical row once, not every
  // continuation, so fragment-local tracks must derive from retained content.
  const fragmentInput: TableRowLayoutInput = {
    ...row,
    id: `${row.id}:fragment:${cursor.rowFragmentIndex}`,
    heightPt: null,
    heightRule: 'auto',
    cells: selectedCells.map((cell, index) => ({
      ...cell.input,
      id: `${cell.input.id}:fragment:${cursor.rowFragmentIndex}:${index}`,
    })),
  };
  return {
    selected: {
      input: fragmentInput,
      logicalRowIndex: row.logicalRowIndex,
      fragmentIndex: cursor.rowFragmentIndex,
      ownership: 'source',
      ranges: selectedCells.map((cell) => cell.range),
    },
    next: Object.freeze({
      rowIndex: complete ? cursor.rowIndex + 1 : cursor.rowIndex,
      rowFragmentIndex: complete ? 0 : cursor.rowFragmentIndex + 1,
      cells: complete ? Object.freeze([]) : Object.freeze(selectedCells.map((cell) => cell.next)),
    }),
    complete,
  };
}

function materializeFragment(
  source: RetainedTableAcquisition,
  selected: readonly SelectedRow[],
  context: TableFragmentContext,
): TableFragmentLayout {
  const fragmentInput: TableLayoutInput = {
    ...source.input,
    id: `${source.input.id}:fragment:${context.page.occurrenceId}`,
    rows: selected.map((row) => row.input),
  };
  const laidOut = layoutTable(fragmentInput, context.placement, context.services).layout;
  const rows = laidOut.rows.map((row, rowIndex): TableRowFragmentLayout => {
    const selection = selected[rowIndex]!;
    return Object.freeze({
      ...row,
      logicalRowIndex: selection.logicalRowIndex,
      fragmentIndex: selection.fragmentIndex,
      ownership: selection.ownership,
      occurrenceId: context.page.occurrenceId,
      physicalPageIndex: context.page.physicalPageIndex,
      displayPageNumber: context.page.displayPageNumber,
      cells: Object.freeze(row.cells.map((cell, cellIndex): TableCellFragmentLayout => {
        const verticalMerge = selection.input.cells[cellIndex]?.verticalMerge ?? 'none';
        const sourceCell = selection.input.cells[cellIndex];
        const ownsRestartInFragment = verticalMerge === 'continue' && selected
          .slice(0, rowIndex)
          .some((earlier) => earlier.input.cells.some((candidate) => (
            candidate.verticalMerge === 'restart'
            && candidate.columnStart === sourceCell?.columnStart
            && candidate.columnSpan === sourceCell?.columnSpan
          )));
        return Object.freeze({
          ...cell,
          contentRanges: Object.freeze([...(selection.ranges[cellIndex] ?? [])]),
          ...(verticalMerge === 'continue' && !ownsRestartInFragment
            ? { visualMergeOwnership: 'continuation' as const }
            : {}),
        });
      })),
    });
  });
  const floatingTables = selected.flatMap((selection, rowIndex) => {
    const sourceRow = source.input.rows[selection.logicalRowIndex];
    if (!sourceRow) return [];
    return source.floatingTables.flatMap((occurrence): FloatingTablePlacementLayout[] => {
      const logicalCellIndex = sourceRow.cells.findIndex((cell) => cell.id === occurrence.hostCellId);
      if (logicalCellIndex < 0) return [];
      const ownsAnchorStart = selection.ranges[logicalCellIndex]?.some((range) => (
        range.blockIndex === occurrence.anchorBlockIndex
          && (range.kind === 'whole'
            || (range.kind === 'paragraph' && range.lineStart === 0))
      )) ?? false;
      if (!ownsAnchorStart) return [];

      const selectedCell = selection.input.cells[logicalCellIndex];
      const laidOutCell = rows[rowIndex]?.cells[logicalCellIndex];
      const anchorBlockOffset = selectedCell?.blocks.findIndex((block) => (
        block.sourceBlockIndex === occurrence.anchorBlockIndex
      )) ?? -1;
      const anchorBlock = anchorBlockOffset < 0
        ? undefined : laidOutCell?.blocks[anchorBlockOffset];
      const child = source.nestedById[occurrence.tableId]?.layout;
      if (!laidOutCell || !anchorBlock || !child) {
        throw new Error('Floating table occurrence references missing retained layout data');
      }
      const anchorBounds = Object.freeze({
        xPt: laidOutCell.contentBounds.xPt,
        yPt: laidOutCell.flowBounds.yPt + anchorBlock.offsetPt,
        widthPt: anchorBlock.layout.flowBounds.widthPt,
        heightPt: anchorBlock.layout.flowBounds.heightPt,
      });
      return [Object.freeze({
        kind: 'floating-table-placement' as const,
        occurrenceId: [
          context.page.occurrenceId,
          occurrence.hostCellId,
          occurrence.sourceBlockIndex,
          occurrence.tableId,
        ].join(':'),
        ownership: selection.ownership,
        physicalPageIndex: context.page.physicalPageIndex,
        displayPageNumber: context.page.displayPageNumber,
        ...occurrence,
        anchorBounds,
        child,
      })];
    });
  });
  // Column measurement is acquisition-owned. A fragment may rebuild row and
  // border geometry, but must retain the one authoritative width vector.
  const clipAtPageEnd = selected.some((row) => row.clipAtPageEnd === true);
  const clippedHeightPt = clipAtPageEnd
    ? Math.min(laidOut.advancePt, context.availableHeightPt)
    : laidOut.advancePt;
  const flowBounds = clipAtPageEnd
    ? { ...laidOut.flowBounds, heightPt: clippedHeightPt }
    : laidOut.flowBounds;
  return Object.freeze({
    ...laidOut,
    flowBounds,
    ...(clipAtPageEnd ? {
      inkBounds: flowBounds,
      clipBounds: flowBounds,
      advancePt: clippedHeightPt,
    } : {}),
    columnWidthsPt: source.layout.columnWidthsPt,
    rows: Object.freeze(rows),
    floatingTables: Object.freeze(floatingTables),
  });
}

export function takeTableFragment(
  source: RetainedTableAcquisition,
  cursor: TableFragmentCursor,
  context: TableFragmentContext,
): TableFragmentResult {
  if (cursor.rowIndex >= source.input.rows.length) {
    return { fragment: null, nextCursor: null, requiresFreshPage: false };
  }

  const selected: SelectedRow[] = [];
  let availablePt = Math.max(0, context.availableHeightPt);
  const headerCount = leadingHeaderCount(source.input);
  if (cursor.rowIndex >= headerCount && cursor.rowIndex > 0 && headerCount > 0) {
    for (let rowIndex = 0; rowIndex < headerCount; rowIndex += 1) {
      const header = rowForOccurrence(
        source,
        source.input.rows[rowIndex]!,
        'repeated-header',
        context,
      );
      const heightPt = paginationRowHeightForOccurrence(source, header, rowIndex, context);
      if (heightPt > availablePt + EPSILON_PT) {
        return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
      }
      selected.push(selectedWholeRow(header, 'repeated-header'));
      availablePt -= heightPt;
    }
  }

  let nextCursor: TableFragmentCursor | null = cursor;
  let rowIndex = cursor.rowIndex;
  while (rowIndex < source.input.rows.length) {
    const ownership: TableFragmentOwnership = 'source';
    const row = rowForOccurrence(source, source.input.rows[rowIndex]!, ownership, context);
    const wholeHeightPt = paginationRowHeightForOccurrence(source, row, rowIndex, context);
    const canTakeWhole = rowIndex !== cursor.rowIndex || cursor.rowFragmentIndex === 0;
    if (canTakeWhole) {
      if (wholeHeightPt <= availablePt + EPSILON_PT) {
        selected.push(selectedWholeRow(row, 'source'));
        availablePt -= wholeHeightPt;
        rowIndex += 1;
        nextCursor = rowIndex < source.input.rows.length
          ? Object.freeze({ rowIndex, rowFragmentIndex: 0, cells: Object.freeze([]) })
          : null;
        continue;
      }
    }

    if (row.cantSplit) {
      const selectedSourceRows = selected.some((item) => item.ownership === 'source');
      if (selectedSourceRows) break;
      const freshHeaderHeightPt = context.availableHeightPt - availablePt;
      const fitsFreshBand = wholeHeightPt + freshHeaderHeightPt
        <= context.freshPageHeightPt + EPSILON_PT;
      if (fitsFreshBand) {
        return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
      }
      if (context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt) {
        return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
      }
      // [MS-OI29500] 2.1.120: Word starts an over-page cantSplit row on the
      // fresh page and clips its overflow instead of synthesizing a continuation.
      if (context.compatibility === 'word'
        && context.availableHeightPt + EPSILON_PT >= context.freshPageHeightPt) {
        selected.push(selectedWholeRow(row, 'source', 0, true));
        nextCursor = rowIndex + 1 < source.input.rows.length
          ? Object.freeze({ rowIndex: rowIndex + 1, rowFragmentIndex: 0, cells: Object.freeze([]) })
          : null;
        break;
      }
      // ECMA-376 §17.4.6 permits a row taller than a full page to continue;
      // only the documented Word compatibility mode clips it.
    }

    // Floating overflow is not defined by §17.4.57. The retained floating
    // adapter preserves the established row-boundary policy: after relocation
    // to a fresh band, one over-band row is emitted once instead of being
    // converted into synthetic line fragments. Ordinary tables keep the
    // specification-backed default split policy above.
    if (context.oversizedRowPolicy === 'atomic'
      && selected.every((item) => item.ownership === 'repeated-header')
      && context.availableHeightPt + EPSILON_PT >= context.freshPageHeightPt
      && wholeHeightPt > context.freshPageHeightPt + EPSILON_PT) {
      selected.push(selectedWholeRow(row, 'source'));
      nextCursor = rowIndex + 1 < source.input.rows.length
        ? Object.freeze({ rowIndex: rowIndex + 1, rowFragmentIndex: 0, cells: Object.freeze([]) })
        : null;
      break;
    }

    const partial = partialRow(source, row, rowIndex === cursor.rowIndex
      ? cursor
      : Object.freeze({ rowIndex, rowFragmentIndex: 0, cells: Object.freeze([]) }), availablePt, context);
    if (partial.selected) {
      selected.push(partial.selected);
      nextCursor = partial.next.rowIndex >= source.input.rows.length ? null : partial.next;
    }
    break;
  }

  const sourceRows = selected.filter((row) => row.ownership === 'source');
  if (sourceRows.length === 0) {
    const canProgressOnFreshPage = context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt;
    return {
      fragment: null,
      nextCursor: cursor,
      requiresFreshPage: canProgressOnFreshPage,
    };
  }
  let fragment = materializeFragment(source, selected, context);
  while (fragment.advancePt > context.availableHeightPt + EPSILON_PT) {
    const last = selected.at(-1);
    const sourceCount = selected.filter((row) => row.ownership === 'source').length;
    const wholeSourceRow = last?.ownership === 'source'
      && last.fragmentIndex === 0
      && last.ranges.every((ranges) => ranges.every((range) => range.kind === 'whole'));
    if (!wholeSourceRow || sourceCount <= 1) break;
    selected.pop();
    nextCursor = Object.freeze({
      rowIndex: last.logicalRowIndex,
      rowFragmentIndex: 0,
      cells: Object.freeze([]),
    });
    fragment = materializeFragment(source, selected, context);
  }
  if (fragment.advancePt > context.availableHeightPt + EPSILON_PT
    && context.availableHeightPt + EPSILON_PT < context.freshPageHeightPt
    && fragment.advancePt <= context.freshPageHeightPt + EPSILON_PT) {
    return { fragment: null, nextCursor: cursor, requiresFreshPage: true };
  }
  return {
    fragment,
    nextCursor,
    requiresFreshPage: false,
  };
}
