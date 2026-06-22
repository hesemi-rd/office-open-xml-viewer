import { describe, it, expect } from 'vitest';
import { resolveColumnWidths } from './renderer.js';
import type { DocTable, DocTableRow, DocTableCell } from './types.js';

// State type taken from resolveColumnWidths' signature so the test does not
// depend on RenderState being exported (mirrors table-row-height.test).
type ColState = Parameters<typeof resolveColumnWidths>[2];

// ECMA-376 §17.4.48 (`<w:tblGrid>`) / §17.4.16 (`<w:gridCol>`) define the
// column widths; per-cell `<w:tcW>` (§17.4.71) is only a PREFERRED width. Word
// bakes its resolved auto-fit widths back into the saved grid, so we size
// columns from the grid (scaled to the table width) and do NOT re-apply tcW
// (see resolveColumnWidths' comment for why this matches Word over the literal
// §17.4.52 algorithm). sample-3's résumé regressed because its rows carry
// single-column `tcW≈30%` each that, under the old tcW-over-grid path,
// flattened an intentionally-unequal grid ([2137, 222, 2430, …] twips) toward
// equal columns, shifting later columns right and changing the wrap.
//
// Empty cells carry no content, so the min-content floor is 0 and `state` is
// never dereferenced (mirrors table-row-height.test's EMPTY_STATE).

const EMPTY_STATE = {} as unknown as ColState;

function cell(colSpan: number, widthPct: number | null): DocTableCell {
  return {
    content: [],
    colSpan,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: null,
    widthPct,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
  } as unknown as DocTableCell;
}

function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as unknown as DocTableRow;
}

/** sample-3-shaped table: an unequal 2-column grid whose single-column cells
 *  each prefer ~50% via `tcW=pct`. */
function table(rows: DocTableRow[]): DocTable {
  return {
    colWidths: [70, 30], // grid: deliberately UNEQUAL
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    // layout omitted ⇒ autofit (the branch under test).
  } as unknown as DocTable;
}

describe('resolveColumnWidths — grid (§17.4.48) governs autofit widths, not per-cell tcW (§17.4.71)', () => {
  it('keeps the (unequal) tblGrid proportions even when cells prefer ~equal tcW=pct', () => {
    // Two rows of single-column cells, each preferring 50% (2500/5000). The grid
    // says 70/30. Word renders 70/30 (the grid); tcW must NOT equalize to 50/50.
    const t = table([
      row([cell(1, 2500), cell(1, 2500)]),
      row([cell(1, 2500), cell(1, 2500)]),
    ]);
    const w = resolveColumnWidths(t, 100, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(70, 5);
    expect(w[1]).toBeCloseTo(30, 5);
  });

  it('a gridSpan cell whose tcW exceeds its grid span does not widen the columns', () => {
    // A 2-col-spanning cell prefers 100% (5000pct) — wider than the grid's full
    // 100 pt. The grid still governs: columns stay 70/30, total stays 100.
    const t = table([row([cell(2, 5000)])]);
    const w = resolveColumnWidths(t, 100, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(70, 5);
    expect(w[1]).toBeCloseTo(30, 5);
  });

  it('scales the grid proportionally when it overflows the available width', () => {
    // Grid sums to 100 pt but only 50 pt is available ⇒ scale by 0.5, preserving
    // the 70/30 proportion.
    const t = table([row([cell(1, 2500), cell(1, 2500)])]);
    const w = resolveColumnWidths(t, 50, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(35, 5);
    expect(w[1]).toBeCloseTo(15, 5);
  });
});
