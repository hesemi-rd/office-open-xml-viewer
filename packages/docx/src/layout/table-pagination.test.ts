import { describe, expect, it } from 'vitest';
import type { RetainedTableAcquisition } from './table-acquisition.js';
import {
  startTableFragmentCursor,
  takeTableFragment,
  type TableFragmentContext,
} from './table-pagination.js';
import { layoutTable } from './table.js';
import type {
  AcquiredParagraphLayoutInput,
  LayoutServices,
  ParagraphLayout,
  TableEdgeInputs,
  TableLayoutInput,
  TableRowLayoutInput,
} from './types.js';
import { layoutParagraph } from './paragraph.js';
import {
  retainTableEnvelope,
  retainedTableEnvelopeFor,
  retainedTableSliceSize,
} from './retained-table-fragments.js';

const noBorders: TableEdgeInputs = {
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
};

function paragraph(id: string, lineHeights: readonly number[]): ParagraphLayout {
  let yPt = 0;
  const lines = lineHeights.map((heightPt, index) => {
    const line = {
      range: { start: index, end: index + 1 },
      bounds: { xPt: 0, yPt, widthPt: 20, heightPt },
      baselinePt: yPt + heightPt * 0.8,
      advancePt: heightPt,
      placements: [],
    };
    yPt += heightPt;
    return line;
  });
  return layoutParagraph({
    kind: 'paragraph', id,
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'cell', ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt: 0, widthPt: 20, heightPt: yPt },
    inkBounds: { xPt: 0, yPt: 0, widthPt: 20, heightPt: yPt },
    spacing: { beforePt: 0, afterPt: 0 },
    contextualSpacing: false,
    lines,
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  } satisfies AcquiredParagraphLayoutInput);
}

function row(
  logicalRowIndex: number,
  heightPt: number,
  options: {
    cantSplit?: boolean;
    repeatedHeader?: boolean;
    heightRule?: 'auto' | 'atLeast' | 'exact';
    paragraph?: ParagraphLayout;
    verticalMerge?: 'none' | 'restart' | 'continue';
  } = {},
): TableRowLayoutInput {
  const p = options.paragraph ?? paragraph(`p-${logicalRowIndex}`, [heightPt]);
  return {
    id: `row-${logicalRowIndex}`,
    source: { story: 'body', storyInstance: 'body', path: [0, logicalRowIndex] },
    logicalRowIndex,
    cantSplit: options.cantSplit ?? false,
    heightPt: options.heightRule === 'auto' || options.heightRule === undefined ? null : heightPt,
    heightRule: options.heightRule ?? 'auto',
    cellSpacingPt: 0,
    exceptionBorders: null,
    alignment: 'left', indentPt: 0,
    repeatedHeader: options.repeatedHeader ?? false,
    cells: [{
      id: `cell-${logicalRowIndex}`,
      source: { story: 'body', storyInstance: 'body', path: [0, logicalRowIndex, 0] },
      columnStart: 0, columnSpan: 1,
      verticalMerge: options.verticalMerge ?? 'none',
      margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      vAlign: 'top', borders: noBorders,
      blocks: options.verticalMerge === 'continue' ? [] : [{
        layout: p,
        sourceBlockIndex: 0,
      }],
    }],
  };
}

function acquisition(rows: readonly TableRowLayoutInput[]): RetainedTableAcquisition {
  const input: TableLayoutInput = {
    kind: 'table', id: 'table-0',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: true,
    alignment: 'left', indentPt: 0, bidiVisual: false,
    columnWidthsPt: [100], borders: noBorders, rows,
  };
  const placement = {
    container: {
      id: 'body', kind: 'body' as const,
      bounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 500 },
    },
    cursor: { xPt: 10, yPt: 20 },
    availableBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 500 },
  };
  return Object.freeze({
    input,
    layout: layoutTable(input, placement, {} as LayoutServices).layout,
    nestedById: Object.freeze({}),
    floatingTables: Object.freeze([]),
  });
}

function withNestedFloatingTable(
  source: RetainedTableAcquisition,
  nested: RetainedTableAcquisition,
  rowIndex = 0,
): RetainedTableAcquisition {
  const sourceRow = source.input.rows[rowIndex]!;
  const hostCellId = sourceRow.cells[0]!.id;
  // The floating table occupied source block 0 before acquisition removed it;
  // retain the following paragraph's original source identity as block 1.
  const input = {
    ...source.input,
    rows: source.input.rows.map((item, index) => index === rowIndex ? {
      ...item,
      cells: item.cells.map((cell, cellIndex) => cellIndex === 0 ? {
        ...cell,
        blocks: cell.blocks.map((block) => ({ ...block, sourceBlockIndex: 1 })),
      } : cell),
    } : item),
  };
  return Object.freeze({
    ...source,
    input,
    nestedById: Object.freeze({ [nested.layout.id]: nested }),
    floatingTables: Object.freeze([Object.freeze({
      hostCellId,
      sourceBlockIndex: 0,
      anchorBlockIndex: 1,
      tableId: nested.layout.id,
      overlap: 'never' as const,
      positioning: Object.freeze({
        leftFromTextPt: 1,
        rightFromTextPt: 2,
        topFromTextPt: 3,
        bottomFromTextPt: 4,
        horzAnchor: 'text',
        horzSpecified: true,
        vertAnchor: 'text',
        xPt: 5,
        yPt: 6,
      }),
    })]),
  });
}

function withFloatingTableAfterLeadingBlock(
  nested: RetainedTableAcquisition,
): RetainedTableAcquisition {
  const baseRow = row(0, 60);
  const hostCell = baseRow.cells[0]!;
  const inputRow: TableRowLayoutInput = {
    ...baseRow,
    cells: [{
      ...hostCell,
      blocks: [
        { layout: paragraph('leading-block', [30]), sourceBlockIndex: 0 },
        { layout: paragraph('later-anchor', [30]), sourceBlockIndex: 2 },
      ],
    }],
  };
  const base = acquisition([inputRow]);
  return Object.freeze({
    ...base,
    nestedById: Object.freeze({ [nested.layout.id]: nested }),
    floatingTables: Object.freeze([Object.freeze({
      hostCellId: hostCell.id,
      sourceBlockIndex: 1,
      anchorBlockIndex: 2,
      tableId: nested.layout.id,
      overlap: 'never' as const,
      positioning: Object.freeze({
        leftFromTextPt: 0,
        rightFromTextPt: 0,
        topFromTextPt: 0,
        bottomFromTextPt: 0,
        horzAnchor: 'page',
        horzSpecified: true,
        vertAnchor: 'text',
        xPt: 5,
        yPt: 0,
      }),
    })]),
  });
}

function take(
  source: RetainedTableAcquisition,
  availableHeightPt: number,
  cursor = startTableFragmentCursor(),
  overrides: Partial<TableFragmentContext> = {},
) {
  return takeTableFragment(source, cursor, {
    availableHeightPt,
    freshPageHeightPt: 100,
    placement: {
      container: {
        id: 'body', kind: 'body',
        bounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: availableHeightPt },
      },
      cursor: { xPt: 10, yPt: 20 },
      availableBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: availableHeightPt },
    },
    services: {} as LayoutServices,
    compatibility: 'word',
    page: { physicalPageIndex: 0, displayPageNumber: 1, occurrenceId: 'page-0' },
    ...overrides,
  });
}

describe('retained table pagination', () => {
  it('keeps retained envelope geometry immutable and isolated by weak object identity', () => {
    const first = {};
    const second = {};
    const layout = acquisition([row(0, 20)]).layout;
    const placement = { fragment: layout, xPt: 5, yPt: 7, widthPt: 100 };

    retainTableEnvelope(first, placement);

    placement.xPt = 99;
    expect(retainedTableEnvelopeFor(first)).toEqual({
      fragment: layout, xPt: 5, yPt: 7, widthPt: 100,
    });
    expect(Object.isFrozen(retainedTableEnvelopeFor(first))).toBe(true);
    expect(retainedTableEnvelopeFor(second)).toBeUndefined();
    expect(retainedTableSliceSize(first, 2)).toEqual({ widthPx: 200, heightPx: 40 });
  });

  it.each(['page', 'margin'] as const)(
    'reflows a %s-relative nested float before committing row selection',
    (anchor) => {
    const nested = acquisition([row(0, 30)]);
    const initial = acquisition([
      row(0, 20, {
        cantSplit: true,
        paragraph: paragraph('page-anchor', [20]),
      }),
    ]);
    const source = withNestedFloatingTable(initial, nested);
    const positioning = {
      ...source.floatingTables[0]!.positioning,
      horzAnchor: anchor,
      vertAnchor: anchor,
      xPt: 10,
      yPt: 20,
    };
    const pageRelative = Object.freeze({
      ...source,
      floatingTables: Object.freeze([Object.freeze({
        ...source.floatingTables[0]!, positioning: Object.freeze(positioning),
      })]),
    });
    const committed: unknown[] = [];
    const wrapped = paragraph('page-anchor-wrapped', [20, 20, 20]);
    const finalFloatContext = {
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 100 },
        margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 80 },
        column: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 80 },
      },
      floatingTableRegistry: {
        coordinateSpace: 'logical-page-points',
        flowDomainId: pageRelative.input.flowDomainId,
        entries: Object.freeze([]),
        nextParagraphId: 0,
      },
      finalPlacementTranslationPt: { xPt: 0, yPt: 0 },
      reacquirePageDependentBlock: (request: {
        acquired: ParagraphLayout | TableLayoutInput;
        floatingTableExclusions?: readonly unknown[];
      }) => request.floatingTableExclusions?.length ? wrapped : request.acquired,
    } as unknown as Partial<TableFragmentContext>;

    const rejected = take(pageRelative, 50, startTableFragmentCursor(), finalFloatContext);

    expect(rejected.fragment).toBeNull();
    expect(rejected.requiresFreshPage).toBe(true);
    committed.push(...(rejected.floatingTablePlacements ?? []));
    expect(committed).toEqual([]);

    const accepted = take(pageRelative, 100, startTableFragmentCursor(), finalFloatContext);
    expect(accepted.fragment?.advancePt).toBe(60);
    expect(accepted.fragment?.rows[0]?.cells[0]?.blocks[0]?.layout).toMatchObject({
      id: wrapped.id,
      lines: { length: 3 },
    });
    committed.push(...(accepted.floatingTablePlacements ?? []));
    expect(committed).toHaveLength(1);
    expect(accepted.fragment?.floatingTables).toEqual([]);
    expect(accepted.fragment?.resolvedFloatingTables[0]).toBe(
      accepted.floatingTablePlacements?.[0],
    );
    expect(Object.isFrozen(accepted.floatingTablePlacements?.[0]?.bounds)).toBe(true);
    expect(JSON.parse(JSON.stringify(accepted.fragment))).toEqual(accepted.fragment);
    },
  );

  it('does not resolve or commit a final-frame float on a paragraph continuation', () => {
    const nested = acquisition([row(0, 10)]);
    const source = withNestedFloatingTable(acquisition([
      row(0, 40, { paragraph: paragraph('continued-anchor', [20, 20]) }),
    ]), nested);
    const pageRelative = Object.freeze({
      ...source,
      floatingTables: Object.freeze([Object.freeze({
        ...source.floatingTables[0]!,
        positioning: Object.freeze({
          ...source.floatingTables[0]!.positioning,
          horzAnchor: 'page', vertAnchor: 'page', xPt: 10, yPt: 20,
        }),
      })]),
    });
    const context = {
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 100 },
        margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 80 },
        column: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 80 },
      },
      floatingTableRegistry: {
        coordinateSpace: 'logical-page-points',
        flowDomainId: pageRelative.input.flowDomainId,
        entries: Object.freeze([]), nextParagraphId: 0,
      },
      finalPlacementTranslationPt: { xPt: 0, yPt: 0 },
      reacquirePageDependentBlock: (request: { acquired: ParagraphLayout }) => request.acquired,
    } as unknown as Partial<TableFragmentContext>;

    const first = take(pageRelative, 20, startTableFragmentCursor(), context);
    const continuation = take(pageRelative, 20, first.nextCursor!, context);

    expect(first.floatingTablePlacements).toHaveLength(1);
    expect(continuation.floatingTablePlacements).toEqual([]);
    expect(continuation.fragment?.floatingTables).toEqual([]);
    expect(continuation.fragment?.resolvedFloatingTables).toEqual([]);
  });

  it('defers a final-frame float until the fragment that owns a later anchor start', () => {
    const source = withFloatingTableAfterLeadingBlock(acquisition([row(0, 10)]));
    const context = {
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 100 },
        margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 80 },
        column: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 80 },
      },
      floatingTableRegistry: {
        coordinateSpace: 'logical-page-points',
        flowDomainId: source.input.flowDomainId,
        entries: Object.freeze([]), nextParagraphId: 0,
      },
      finalPlacementTranslationPt: { xPt: 0, yPt: 0 },
      reacquirePageDependentBlock: (request: { acquired: ParagraphLayout }) => request.acquired,
    } as unknown as Partial<TableFragmentContext>;

    const beforeAnchor = take(source, 30, startTableFragmentCursor(), context);

    expect(beforeAnchor.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 },
    ]);
    expect(beforeAnchor.floatingTablePlacements).toEqual([]);
    expect(beforeAnchor.floatingTableRegistryDelta).toEqual({
      coordinateSpace: 'logical-page-points',
      flowDomainId: source.input.flowDomainId,
      baseNextParagraphId: 0,
      nextParagraphId: 0,
      entries: [],
    });
    expect(beforeAnchor.fragment?.resolvedFloatingTables).toEqual([]);

    const atAnchor = take(source, 30, beforeAnchor.nextCursor!, context);
    expect(atAnchor.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 2, lineStart: 0, lineEnd: 1 },
    ]);
    expect(atAnchor.floatingTablePlacements).toHaveLength(1);
    expect(atAnchor.floatingTablePlacements?.[0]?.source.anchorBounds.yPt).toBe(20);
    expect(structuredClone(atAnchor.floatingTableRegistryDelta)).toEqual(
      atAnchor.floatingTableRegistryDelta,
    );
    expect(atAnchor.floatingTableRegistryDelta).toMatchObject({
      coordinateSpace: 'logical-page-points',
      flowDomainId: source.input.flowDomainId,
      baseNextParagraphId: 0,
      nextParagraphId: 1,
      entries: [{ paragraphId: 0 }],
    });
  });

  it('recomputes later mixed-axis placements after earlier anchor reflow', () => {
    const nested = acquisition([row(0, 10)]);
    const baseRow = row(0, 40);
    const host = baseRow.cells[0]!;
    const inputRow: TableRowLayoutInput = {
      ...baseRow,
      cells: [{
        ...host,
        blocks: [
          { layout: paragraph('first-anchor', [20]), sourceBlockIndex: 1 },
          { layout: paragraph('second-anchor', [20]), sourceBlockIndex: 3 },
        ],
      }],
    };
    const base = acquisition([inputRow]);
    const positioning = Object.freeze({
      leftFromTextPt: 0, rightFromTextPt: 0, topFromTextPt: 0, bottomFromTextPt: 0,
      horzAnchor: 'page', horzSpecified: true, vertAnchor: 'text', xPt: 10, yPt: 0,
    });
    const source: RetainedTableAcquisition = Object.freeze({
      ...base,
      nestedById: Object.freeze({ [nested.layout.id]: nested }),
      floatingTables: Object.freeze([0, 2].map((sourceBlockIndex, index) => Object.freeze({
        hostCellId: host.id,
        sourceBlockIndex,
        anchorBlockIndex: index === 0 ? 1 : 3,
        tableId: nested.layout.id,
        overlap: 'never' as const,
        positioning,
      }))),
    });
    const context = {
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 120 },
        margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 100 },
        column: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 100 },
      },
      floatingTableRegistry: {
        coordinateSpace: 'logical-page-points', flowDomainId: 'logical-page:0',
        entries: Object.freeze([]), nextParagraphId: 0,
      },
      reacquirePageDependentBlock: (request: { sourceBlockIndex: number; acquired: ParagraphLayout }) => (
        request.sourceBlockIndex === 1 ? paragraph('first-anchor-wrapped', [20, 20]) : request.acquired
      ),
    } as unknown as Partial<TableFragmentContext>;

    const result = take(source, 100, startTableFragmentCursor(), context);

    expect(result.floatingTablePlacements).toHaveLength(2);
    expect(result.floatingTablePlacements?.map((item) => item.source.anchorBounds.yPt)).toEqual([
      20, 60,
    ]);
    expect(result.floatingTablePlacements?.[1]?.yPt).toBe(60);
    expect(result.floatingTableRegistryDelta?.entries.map((entry) => entry.paragraphId)).toEqual([
      0, 1,
    ]);
  });

  it('emits the largest fitting row prefix and preserves one column authority', () => {
    const source = acquisition([row(0, 30), row(1, 30), row(2, 30)]);

    const first = take(source, 65);
    expect(first.requiresFreshPage).toBe(false);
    expect(first.fragment?.rows.map((item) => item.logicalRowIndex)).toEqual([0, 1]);
    expect(first.fragment?.columnWidthsPt).toBe(source.layout.columnWidthsPt);
    expect(first.fragment?.advancePt).toBe(60);
    expect(first.nextCursor).toMatchObject({ rowIndex: 2, rowFragmentIndex: 0 });

    const second = take(source, 65, first.nextCursor!);
    expect(second.fragment?.rows.map((item) => item.logicalRowIndex)).toEqual([2]);
    expect(second.nextCursor).toBeNull();
  });

  it('requests a fresh page for an unbreakable row that fits the fresh band', () => {
    const source = acquisition([row(0, 70, { cantSplit: true })]);
    const cursor = startTableFragmentCursor();

    const result = take(source, 40, cursor);

    expect(result.fragment).toBeNull();
    expect(result.requiresFreshPage).toBe(true);
    expect(result.nextCursor).toEqual(cursor);
  });

  it('moves a fully retained exact-height row instead of discarding its authored box', () => {
    const source = acquisition([row(0, 90, {
      heightRule: 'exact',
      paragraph: paragraph('complete-content', [20]),
    })]);
    const cursor = startTableFragmentCursor();

    const constrained = take(source, 80, cursor);

    expect(constrained.fragment).toBeNull();
    expect(constrained.requiresFreshPage).toBe(true);
    expect(constrained.nextCursor).toEqual(cursor);

    const fresh = take(source, 100, cursor);
    expect(fresh.fragment?.advancePt).toBe(90);
    expect(fresh.nextCursor).toBeNull();
  });

  it('keeps the largest fitting exact-row prefix before considering a fresh page', () => {
    const source = acquisition([
      row(0, 50, { heightRule: 'exact', paragraph: paragraph('first', [20]) }),
      row(1, 50, { heightRule: 'exact', paragraph: paragraph('second', [20]) }),
    ]);

    const result = take(source, 90);

    expect(result.requiresFreshPage).toBe(false);
    expect(result.fragment?.rows.map((item) => item.logicalRowIndex)).toEqual([0]);
    expect(result.fragment?.advancePt).toBe(50);
    expect(result.nextCursor).toMatchObject({ rowIndex: 1, rowFragmentIndex: 0 });
  });

  it('splits an over-page cantSplit row in standard mode', () => {
    const source = acquisition([row(0, 120, {
      cantSplit: true,
      paragraph: paragraph('standard-over-page', [40, 40, 40]),
    })]);

    const result = take(source, 100, startTableFragmentCursor(), {
      compatibility: 'standard',
    });

    expect(result.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 2 },
    ]);
    expect(result.nextCursor).toMatchObject({ rowIndex: 0, rowFragmentIndex: 1 });
  });

  it('clips an over-page cantSplit row to the fresh Word page band without continuation', () => {
    const source = acquisition([row(0, 120, {
      cantSplit: true,
      paragraph: paragraph('word-over-page', [40, 40, 40]),
    })]);

    const result = take(source, 100);

    expect(result.fragment?.advancePt).toBe(100);
    expect(result.fragment?.flowBounds.heightPt).toBe(100);
    expect(result.fragment?.clipBounds?.heightPt).toBe(100);
    expect(result.nextCursor).toBeNull();
  });

  it('splits a paragraph by retained line boundaries without reacquiring text', () => {
    const retained = paragraph('multi-line', [20, 20, 20]);
    const source = acquisition([row(0, 60, { paragraph: retained })]);

    const first = take(source, 45);
    const firstParagraph = first.fragment?.rows[0]?.cells[0]?.blocks[0]?.layout;
    expect(firstParagraph?.kind).toBe('paragraph');
    expect(firstParagraph && firstParagraph.kind === 'paragraph'
      ? firstParagraph.continuation : null).toEqual({
      lineStart: 0, lineEnd: 2, continuesFromPrevious: false, continuesOnNext: true,
    });
    expect(first.fragment?.rows[0]?.fragmentIndex).toBe(0);
    expect(first.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 2 },
    ]);

    const second = take(source, 45, first.nextCursor!);
    const secondParagraph = second.fragment?.rows[0]?.cells[0]?.blocks[0]?.layout;
    expect(secondParagraph && secondParagraph.kind === 'paragraph'
      ? secondParagraph.continuation : null).toEqual({
      lineStart: 2, lineEnd: 3, continuesFromPrevious: true, continuesOnNext: false,
    });
    expect(second.fragment?.rows[0]?.fragmentIndex).toBe(1);
    expect(second.nextCursor).toBeNull();
    expect(retained.continuation).toBeUndefined();
  });

  it('repeats only the leading header prefix without consuming source ownership twice', () => {
    const source = acquisition([
      row(0, 20, { repeatedHeader: true }),
      row(1, 20, { repeatedHeader: true }),
      row(2, 40),
      row(3, 40, { repeatedHeader: true }),
    ]);
    const first = take(source, 80);
    const second = take(source, 80, first.nextCursor!);

    expect(first.fragment?.rows.map((item) => [item.logicalRowIndex, item.ownership])).toEqual([
      [0, 'source'], [1, 'source'], [2, 'source'],
    ]);
    expect(second.fragment?.rows.map((item) => [item.logicalRowIndex, item.ownership])).toEqual([
      [0, 'repeated-header'], [1, 'repeated-header'], [3, 'source'],
    ]);
    expect(second.nextCursor).toBeNull();
  });

  it('keeps vMerge source roles immutable when a page boundary cuts the span', () => {
    const source = acquisition([
      row(0, 60, { verticalMerge: 'restart' }),
      row(1, 60, { verticalMerge: 'continue' }),
    ]);

    const first = take(source, 60);
    const second = take(source, 60, first.nextCursor!);

    expect(first.fragment?.rows[0]?.cells[0]?.verticalMerge).toBe('restart');
    expect(second.fragment?.rows[0]?.cells[0]?.verticalMerge).toBe('continue');
    expect(second.fragment?.rows[0]?.cells[0]?.visualMergeOwnership).toBe('continuation');
    expect(source.input.rows[1]?.cells[0]?.verticalMerge).toBe('continue');
  });

  it('does not treat exact height or repeated-header as implicit cantSplit', () => {
    const exact = paragraph('exact', [30, 30]);
    const source = acquisition([
      row(0, 60, { heightRule: 'exact', paragraph: exact, repeatedHeader: true }),
    ]);

    const first = take(source, 35);

    expect(first.requiresFreshPage).toBe(false);
    expect(first.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 },
    ]);
  });

  it('reacquires only page-dependent blocks with stable source indices per occurrence', () => {
    const header = row(0, 20, { repeatedHeader: true });
    const dependentHeader = {
      ...header,
      cells: header.cells.map((cell) => ({
        ...cell,
        blocks: cell.blocks.map((block) => ({ ...block, pageDependent: true })),
      })),
    };
    const source = acquisition([dependentHeader, row(1, 60), row(2, 60)]);
    const calls: Array<{
      rowIndex: number;
      cellIndex: number;
      blockIndex: number;
      ownership: string;
      occurrenceId: string;
    }> = [];
    const reacquire: NonNullable<TableFragmentContext['reacquirePageDependentBlock']> = (request) => {
      calls.push({
        rowIndex: request.logicalRowIndex,
        cellIndex: request.logicalCellIndex,
        blockIndex: request.sourceBlockIndex,
        ownership: request.ownership,
        occurrenceId: request.page.occurrenceId,
      });
      return paragraph(`page-${request.page.displayPageNumber}`, [20]);
    };

    const first = take(source, 80, startTableFragmentCursor(), {
      page: { physicalPageIndex: 0, displayPageNumber: 9, occurrenceId: 'page-9' },
      reacquirePageDependentBlock: reacquire,
    });
    const second = take(source, 80, first.nextCursor!, {
      page: { physicalPageIndex: 1, displayPageNumber: 10, occurrenceId: 'page-10' },
      reacquirePageDependentBlock: reacquire,
    });

    expect(calls).toEqual([
      { rowIndex: 0, cellIndex: 0, blockIndex: 0, ownership: 'source', occurrenceId: 'page-9' },
      { rowIndex: 0, cellIndex: 0, blockIndex: 0, ownership: 'repeated-header', occurrenceId: 'page-10' },
    ]);
    expect(first.fragment?.rows[0]?.cells[0]?.blocks[0]?.layout.id).toBe('page-9');
    expect(second.fragment?.rows[0]?.cells[0]?.blocks[0]?.layout.id).toBe('page-10');
    expect(second.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'whole', blockIndex: 0 },
    ]);
    expect(first.fragment?.rows[0]).toMatchObject({
      occurrenceId: 'page-9',
      physicalPageIndex: 0,
      displayPageNumber: 9,
    });
    expect(second.fragment?.rows[0]).toMatchObject({
      occurrenceId: 'page-10',
      physicalPageIndex: 1,
      displayPageNumber: 10,
      ownership: 'repeated-header',
    });
  });

  it('uses reacquired header geometry in the destination-page fit decision', () => {
    const header = row(0, 20, { repeatedHeader: true });
    const dependentHeader = {
      ...header,
      cells: header.cells.map((cell) => ({
        ...cell,
        blocks: cell.blocks.map((block) => ({ ...block, pageDependent: true })),
      })),
    };
    const body = row(1, 60, { paragraph: paragraph('body-lines', [20, 40]) });
    const source = acquisition([dependentHeader, body]);
    const first = take(source, 20);

    const second = take(source, 80, first.nextCursor!, {
      page: { physicalPageIndex: 1, displayPageNumber: 10, occurrenceId: 'page-10' },
      reacquirePageDependentBlock: () => paragraph('wide-page-number', [40]),
    });

    expect(second.fragment?.advancePt).toBe(60);
    expect(second.fragment?.rows.map((item) => [item.logicalRowIndex, item.advancePt])).toEqual([
      [0, 40], [1, 20],
    ]);
    expect(second.nextCursor).toMatchObject({ rowIndex: 1, rowFragmentIndex: 1 });
  });

  it('continues a nested retained table with its own immutable cursor', () => {
    const nested = acquisition([row(0, 30), row(1, 30)]);
    const outerRow = row(0, 60);
    const outerInput: TableLayoutInput = {
      ...acquisition([outerRow]).input,
      rows: [{
        ...outerRow,
        cells: [{
          ...outerRow.cells[0]!,
          blocks: [{ layout: nested.layout, sourceBlockIndex: 0 }],
        }],
      }],
    };
    const outer = acquisition(outerInput.rows);
    const source: RetainedTableAcquisition = Object.freeze({
      ...outer,
      input: outerInput,
      nestedById: Object.freeze({ [nested.layout.id]: nested }),
    });

    const first = take(source, 35);
    const second = take(source, 35, first.nextCursor!);

    expect(first.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'nested-table', blockIndex: 0, childFragmentIndex: 0 },
    ]);
    expect(second.fragment?.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'nested-table', blockIndex: 0, childFragmentIndex: 1 },
    ]);
    expect(second.nextCursor).toBeNull();
  });

  it('projects a nested floating table only onto the split-row fragment owning its anchor start', () => {
    const nested = acquisition([row(0, 10)]);
    const source = withNestedFloatingTable(acquisition([
      row(0, 40, { paragraph: paragraph('split-anchor', [20, 20]) }),
    ]), nested);

    const first = take(source, 25, startTableFragmentCursor(), {
      page: { physicalPageIndex: 0, displayPageNumber: 1, occurrenceId: 'page-0' },
    });
    const second = take(source, 25, first.nextCursor!, {
      page: { physicalPageIndex: 1, displayPageNumber: 2, occurrenceId: 'page-1' },
    });

    expect(first.fragment?.floatingTables).toHaveLength(1);
    expect(second.fragment?.floatingTables).toEqual([]);
    const placement = first.fragment!.floatingTables[0]!;
    const anchorCell = first.fragment!.rows[0]!.cells[0]!;
    const anchorBlock = anchorCell.blocks[0]!;
    expect(placement).toMatchObject({
      kind: 'floating-table-placement',
      occurrenceId: `page-0:${source.input.rows[0]!.cells[0]!.id}:0:${nested.layout.id}`,
      ownership: 'source',
      physicalPageIndex: 0,
      displayPageNumber: 1,
      hostCellId: source.input.rows[0]!.cells[0]!.id,
      sourceBlockIndex: 0,
      anchorBlockIndex: 1,
      tableId: nested.layout.id,
      overlap: 'never',
      positioning: source.floatingTables[0]!.positioning,
      anchorBounds: {
        xPt: anchorCell.contentBounds.xPt,
        yPt: anchorCell.flowBounds.yPt + anchorBlock.offsetPt,
        widthPt: anchorBlock.layout.flowBounds.widthPt,
        heightPt: anchorBlock.layout.flowBounds.heightPt,
      },
    });
    expect(placement.child).toBe(nested.layout);
  });

  it('creates one page-local floating occurrence for each repeated-header occurrence', () => {
    const nested = acquisition([row(0, 10)]);
    const header = row(0, 20, { repeatedHeader: true });
    const retained = withNestedFloatingTable(acquisition([
      header,
      row(1, 60),
      row(2, 60),
    ]), nested);
    const source = Object.freeze({
      ...retained,
      floatingTables: Object.freeze(retained.floatingTables.map((placement) => Object.freeze({
        ...placement,
        positioning: Object.freeze({
          ...placement.positioning,
          horzAnchor: 'page', vertAnchor: 'page', xPt: 10, yPt: 20,
        }),
      }))),
    });
    const finalContext = (pageIndex: number) => ({
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 100 },
        margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 80 },
        column: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 80 },
      },
      floatingTableRegistry: {
        coordinateSpace: 'logical-page-points',
        flowDomainId: `logical-page:${pageIndex}`,
        entries: Object.freeze([]), nextParagraphId: 0,
      },
      reacquirePageDependentBlock: (request: { acquired: ParagraphLayout }) => request.acquired,
    }) as unknown as Partial<TableFragmentContext>;

    const first = take(source, 80, startTableFragmentCursor(), {
      page: { physicalPageIndex: 0, displayPageNumber: 9, occurrenceId: 'page-9' },
      ...finalContext(0),
    });
    const second = take(source, 80, first.nextCursor!, {
      page: { physicalPageIndex: 1, displayPageNumber: 10, occurrenceId: 'page-10' },
      ...finalContext(1),
    });

    expect(first.fragment?.resolvedFloatingTables).toHaveLength(1);
    expect(second.fragment?.resolvedFloatingTables).toHaveLength(1);
    expect(first.fragment?.resolvedFloatingTables[0]?.source).toMatchObject({
      ownership: 'source',
      occurrenceId: expect.stringContaining('page-9:'),
      physicalPageIndex: 0,
      displayPageNumber: 9,
    });
    expect(second.fragment?.resolvedFloatingTables[0]?.source).toMatchObject({
      ownership: 'repeated-header',
      occurrenceId: expect.stringContaining('page-10:'),
      physicalPageIndex: 1,
      displayPageNumber: 10,
    });
    expect(first.fragment?.resolvedFloatingTables[0]?.child).toBe(nested.layout);
    expect(second.fragment?.resolvedFloatingTables[0]?.child).toBe(nested.layout);
    expect(first.floatingTableRegistryDelta?.entries).toHaveLength(1);
    expect(second.floatingTableRegistryDelta?.entries).toHaveLength(1);
  });

});
