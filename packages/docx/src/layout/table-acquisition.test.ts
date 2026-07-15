import { describe, expect, it } from 'vitest';
import type { DocParagraph, DocTable } from '../types.js';
import { tableColumnLayoutInput } from '../parser-model.js';
import { acquireRetainedTableLayout } from './table-acquisition.js';
import type {
  LayoutServices,
  ParagraphLayout,
  TableFormatInput,
} from './types.js';

const noBorders = Object.freeze({
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
});

function retainedParagraph(widthPt: number): ParagraphLayout {
  const bounds = { xPt: 0, yPt: 0, widthPt, heightPt: 8 };
  return {
    kind: 'paragraph', id: 'paragraph',
    source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
    flowDomainId: 'table-cell', ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: 8,
    alignment: 'left', bidi: false,
    spacing: { beforePt: 0, afterPt: 0 },
    borders: [], lines: [],
  } as unknown as ParagraphLayout;
}

describe('retained table acquisition', () => {
  it('includes cell-spacing in autofit intrinsic grid constraints', () => {
    const table = {
      colWidths: [50, 50],
      rows: [{
        gridBefore: 0, gridAfter: 0,
        cells: [0, 1].map(() => ({
          content: [], colSpan: 1, vMerge: null, borders: noBorders,
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        })),
        __tableRowLayout: {
          height: null, justification: null, beforeWidth: null, afterWidth: null,
          cellSpacing: { kind: 'dxa', value: '80' }, exception: null,
        },
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left', layout: 'autofit',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [{ width: '1000' }, { width: '1000' }],
          requiredColumnCount: 2,
        },
        preferredWidth: null, layout: { kind: 'autofit' },
        cellSpacing: null, cellMargins: null,
      },
    } as unknown as DocTable;

    const input = tableColumnLayoutInput(
      table,
      100,
      () => ({ minWidthPt: 10, maxWidthPt: 20 }),
    );

    expect(input.rows[0]?.cells.map((cell) => ({
      min: cell.minContentWidthPt,
      max: cell.maxContentWidthPt,
    }))).toEqual([{ min: 16, max: 26 }, { min: 16, max: 26 }]);
  });

  it('acquires cell children at the final width after cell spacing and margins', () => {
    const paragraph = { type: 'paragraph' } as unknown as DocParagraph;
    const table = {
      rows: [{
        cells: [
          { content: [paragraph], colSpan: 1, vMerge: null, borders: noBorders,
            background: null, vAlign: 'top' },
          { content: [paragraph], colSpan: 1, vMerge: null, borders: noBorders,
            background: null, vAlign: 'top' },
        ],
        gridBefore: 0, gridAfter: 0, isHeader: false,
      }],
      borders: noBorders,
      cellMarginTop: 0, cellMarginRight: 1, cellMarginBottom: 0, cellMarginLeft: 1,
      jc: 'left', bidiVisual: false,
    } as unknown as DocTable;
    const format: TableFormatInput = {
      firstRowException: null,
      rows: [{
        height: null,
        cellSpacingPt: 4,
        justification: null,
        exception: null,
        cells: [
          { marginsPt: { top: 0, right: 1, bottom: 0, left: 1 } },
          { marginsPt: { top: 0, right: 1, bottom: 0, left: 1 } },
        ],
      }],
    };
    const acquiredWidths: number[] = [];

    const layout = acquireRetainedTableLayout(
      table,
      [50, 50],
      100,
      { yPt: 0 },
      [0],
      {
        layoutServices: () => ({}) as LayoutServices,
        tableFormat: () => format,
        resolveColumns: () => [],
        createCellState: (state) => state,
        acquireParagraph: (_state, _paragraph, widthPt) => {
          acquiredWidths.push(widthPt);
          return retainedParagraph(widthPt);
        },
        advanceState: () => {},
      },
    );

    // Each outer cell owns one full outer spacing inset and half of the shared
    // inner gap: 50 - (4 + 2) - (1 + 1) = 42pt.
    expect(acquiredWidths).toEqual([42, 42]);
    expect(layout.rows[0]?.cells.map((cell) => cell.contentBounds.widthPt)).toEqual([42, 42]);
  });
});
