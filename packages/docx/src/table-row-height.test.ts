import { describe, it, expect } from 'vitest';
import { calculateRowHeight } from './renderer.js';
import { resolveSingleRowHeight, resolveTableRowHeights } from './table-geometry.js';
import type { DocTable, DocTableRow, DocTableCell } from './types.js';

// ECMA-376 §17.4.80 (trHeight) + §17.18.37 (ST_HeightRule): the @hRule attribute
// decides how w:trHeight/@val constrains the row height.
//   exact   — height is exactly @val (overflow clipped).
//   atLeast — @val is a lower bound; content can expand the row.
//   auto    — per the §17.4.80 literal, @val is IGNORED ("no predetermined
//             minimum or maximum size", advisory layout cache only). Word's
//             output PDFs, however, treat @val as a LOWER BOUND (same as
//             atLeast) when hRule is omitted and @val is present — e.g.
//             sample-11.docx's December 2007 calendar emits trHeight w:val=576
//             (no hRule, spec default = auto) on its date rows and Word renders
//             each such row at exactly 576 / 20 = 28.8 pt, matching @val as a
//             floor (pdftotext -bbox; the larger per-week cadence is that
//             28.8 pt date row plus an unmarked auto spacer row, not one @val
//             row). XML inspection confirms no other height signal exists. We
//             deliberately deviate from the §17.4.80 literal to match Word's
//             behavior; @val absent still falls back to the implementation-
//             defined minimum.
//
// These tests pin the floor logic with empty cells: with no content, the cell's
// content height is just its (here zero) margins, so the row height is governed
// entirely by the trHeight rule. calculateRowHeight only measures cell content
// when a cell carries any (measureCellElementHeight), so empty cells let us
// exercise the rule branch without a real measuring context — the `state` is
// never dereferenced for empty cells.

const EMPTY_STATE = {} as unknown as Parameters<typeof calculateRowHeight>[4];

function emptyCell(): DocTableCell {
  return {
    content: [],
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: null,
  } as unknown as DocTableCell;
}

function rowWith(rowHeight: number | null, rule: string): DocTableRow {
  return {
    cells: [emptyCell()],
    rowHeight,
    rowHeightRule: rule,
    isHeader: false,
  } as unknown as DocTableRow;
}

function table(): DocTable {
  return {
    colWidths: [100],
    rows: [],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

describe('calculateRowHeight — ST_HeightRule (§17.4.80 / §17.18.37)', () => {
  const COLS = [100];
  const SCALE = 2; // px == 2 * pt, to confirm scale is applied
  const t = table();

  it('exact — height is exactly @val regardless of (here empty) content', () => {
    const h = calculateRowHeight(rowWith(600, 'exact'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(600 * SCALE);
  });

  it('atLeast — @val is a lower bound; empty content cannot shrink below it', () => {
    const h = calculateRowHeight(rowWith(600, 'atLeast'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(600 * SCALE);
  });

  // Word-compatible deviation from the §17.4.80 literal: with hRule omitted
  // (spec default = auto) and @val present, Word honors @val as a lower bound.
  // Ground truth is the Word output PDF — sample-11.docx's December calendar
  // date rows measure exactly trHeight @val 576 / 20 = 28.8 pt. See the
  // resolveSingleRowHeight docstring (table-geometry.ts) for the full rationale.
  it('auto with @val — @val is honored as a lower bound (Word-compatible)', () => {
    const h = calculateRowHeight(rowWith(600, 'auto'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(600 * SCALE);
  });

  it('auto with no @val — falls back to the implementation-defined minimum', () => {
    const h = calculateRowHeight(rowWith(null, 'auto'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(10 * SCALE);
  });

  it('atLeast with no @val — falls back to the minimum (val null ⇒ no floor)', () => {
    const h = calculateRowHeight(rowWith(null, 'atLeast'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(10 * SCALE);
  });

  it('§17.4.15: measures a cell against columns after gridBefore', () => {
    const measuredWidths: number[] = [];
    const row = {
      ...rowWith(null, 'auto'),
      gridBefore: 1,
      gridAfter: 1,
    } as unknown as DocTableRow;

    resolveSingleRowHeight(row, [20, 40, 60], 1, (_cell, width) => {
      measuredWidths.push(width);
      return 10;
    });

    expect(measuredWidths).toEqual([40]);
  });

  it('includes half of each resolved horizontal border in adjacent non-exact row boxes', () => {
    const single = { style: 'single', width: 0.5, color: '#000000' };
    const rows = Array.from({ length: 4 }, () => ({
      ...rowWith(20.4, 'auto'),
      cells: [{
        ...emptyCell(),
        borders: { top: single, bottom: single, left: null, right: null, insideH: null, insideV: null },
      }],
    } as unknown as DocTableRow));
    const bordered = {
      ...table(),
      rows,
    } as DocTable;

    const heights = resolveTableRowHeights(bordered, [100], 1, () => 0);

    // Five 0.5pt boundary rules contribute half at the two outer edges and a
    // full rule across each of the three shared boundaries: 4 × 20.4 + 2.0.
    expect(heights).toEqual([20.9, 20.9, 20.9, 20.9]);
    expect(heights.reduce((sum, height) => sum + height, 0)).toBeCloseTo(83.6, 8);
  });
});
