import { describe, expect, it } from 'vitest';
import { tableColumnLayoutInput } from './parser-model.js';
import type { DocTable } from './types.js';

function emptyBorders() {
  return { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
}

describe('table column acquisition boundary', () => {
  it('normalizes parser-private lexical widths without exposing parser objects to layout', () => {
    const cell = {
      content: [], colSpan: 3, vMerge: null, borders: emptyBorders(),
      background: null, vAlign: 'top', widthPt: null,
      __tableCellLayout: { preferredWidth: { kind: 'pct', value: '2500' }, margins: null },
    };
    const table = {
      colWidths: [36, 0],
      rows: [{
        gridBefore: 1, gridAfter: 1,
        cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
        __tableRowLayout: {
          height: null,
          beforeWidth: { kind: 'pct', value: '15%' },
          afterWidth: { kind: 'dxa', value: '200' },
          cellSpacing: null, exception: null,
        },
      }],
      borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: 'Synthetic',
        grid: {
          authored: true,
          columns: [{ width: '720' }, { width: null }],
          requiredColumnCount: 5,
        },
        preferredWidth: { kind: 'pct', value: '3750' },
        layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    let intrinsicProbeCount = 0;
    const result = tableColumnLayoutInput(table, 200, () => {
      intrinsicProbeCount += 1;
      return { minWidthPt: 12, maxWidthPt: 30 };
    });

    expect(intrinsicProbeCount).toBe(0);

    expect(result).toEqual({
      layout: 'fixed', availableWidthPt: 200,
      gridWidthsPt: [36, 0, 0, 0, 0],
      tablePreferredWidthPt: 150,
      rows: [{
        // wBefore/wAfter percentages use the page text extents (§17.4.85–86),
        // unlike tcW percentages, which remain relative to final table width.
        before: { columnSpan: 1, preferredWidth: { kind: 'dxa', value: 30 } },
        after: { columnSpan: 1, preferredWidth: { kind: 'dxa', value: 10 } },
        cells: [{
          columnStart: 1, columnSpan: 3,
          preferredWidth: { kind: 'pct', value: 0.5 },
          minContentWidthPt: 0, maxContentWidthPt: 0,
        }],
      }],
    });
  });

  it('uses stable public fields only as a compatibility fallback for hand-built tables', () => {
    const cell = {
      content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
      background: null, vAlign: 'top', widthPt: 25, widthPct: null,
    };
    const table = {
      colWidths: [40], rows: [{ cells: [cell], gridBefore: 0, gridAfter: 0 }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'autofit',
      widthPct: 2500,
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({
        layout: 'autofit', gridWidthsPt: [40], tablePreferredWidthPt: 100,
        rows: [{ cells: [{
          columnStart: 0, columnSpan: 1,
          preferredWidth: { kind: 'dxa', value: 25 },
        }] }],
      });
  });

  it('applies first-row tblPrEx fixed layout and width to the whole table in Word mode', () => {
    const table = {
      colWidths: [40],
      rows: [{
        cells: [{
          content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 0, gridAfter: 0,
        __tableRowLayout: {
          height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
          exception: {
            preferredWidth: { kind: 'pct', value: '3000' },
            layout: { kind: 'fixed' }, justification: null, indent: null,
            borders: null, cellMargins: null, cellSpacing: null,
          },
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'autofit',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: '2000' },
        layout: { kind: 'autofit' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({ layout: 'fixed', tablePreferredWidthPt: 120 });
  });

  it.each([
    [{ kind: 'dxa', value: '0' }, 'zero'],
    [{ kind: 'auto', value: '1440' }, 'auto'],
    [{ kind: 'nil', value: '1440' }, 'nil'],
  ] as const)(
    'lets an authored first-row tblPrEx width of %s (%s) shadow the table preferred width',
    (exceptionWidth, _label) => {
      const table = {
        colWidths: [40],
        rows: [{
          cells: [{
            content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
            background: null, vAlign: 'top', widthPt: null,
            __tableCellLayout: { preferredWidth: null, margins: null },
          }],
          gridBefore: 0, gridAfter: 0,
          __tableRowLayout: {
            height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
            exception: {
              preferredWidth: exceptionWidth,
              layout: null, justification: null, indent: null,
              borders: null, cellMargins: null, cellSpacing: null,
            },
          },
        }],
        borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
        cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
        __tableLayout: {
          effectiveStyleId: null,
          grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
          preferredWidth: { kind: 'dxa', value: '2000' },
          layout: { kind: 'fixed' }, cellSpacing: null,
        },
      } as unknown as DocTable;

      expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
        .toMatchObject({ tablePreferredWidthPt: null });
    },
  );

  it('falls back to the parent tblW when a first-row exception omits tblW', () => {
    const table = {
      colWidths: [40],
      rows: [{
        cells: [{
          content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 0, gridAfter: 0,
        __tableRowLayout: {
          height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
          exception: {
            preferredWidth: null,
            layout: { kind: 'fixed' }, justification: null, indent: null,
            borders: null, cellMargins: null, cellSpacing: null,
          },
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'autofit',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: '2000' },
        layout: { kind: 'autofit' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({ layout: 'fixed', tablePreferredWidthPt: 100 });
  });

  it('uses the CT_TblWidth dxa default for an exception width with omitted type', () => {
    const table = {
      colWidths: [40],
      rows: [{
        cells: [{
          content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 0, gridAfter: 0,
        __tableRowLayout: {
          height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
          exception: {
            preferredWidth: { kind: null, value: '1440' },
            layout: null, justification: null, indent: null,
            borders: null, cellMargins: null, cellSpacing: null,
          },
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: '2000' },
        layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({ tablePreferredWidthPt: 72 });
  });

  it('ignores authored gridBefore/gridAfter values which do not fit the table grid', () => {
    const table = {
      colWidths: [20, 40],
      rows: [{
        cells: [{
          content: [], colSpan: 2, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 3, gridAfter: 1,
        __tableRowLayout: {
          height: null, beforeWidth: { kind: 'dxa', value: '100' },
          afterWidth: { kind: 'dxa', value: '100' }, cellSpacing: null, exception: null,
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '400' }, { width: '800' }], requiredColumnCount: 2 },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 })))
      .toMatchObject({
        gridWidthsPt: [20, 40],
        rows: [{ before: null, after: null, cells: [{ columnStart: 0, columnSpan: 2 }] }],
      });
  });
});
