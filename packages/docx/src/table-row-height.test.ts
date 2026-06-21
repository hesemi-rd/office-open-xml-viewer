import { describe, it, expect } from 'vitest';
import { calculateRowHeight } from './renderer.js';
import type { DocTable, DocTableRow, DocTableCell } from './types.js';

// ECMA-376 §17.4.80 (trHeight) + §17.18.37 (ST_HeightRule): the @hRule attribute
// decides how w:trHeight/@val constrains the row height.
//   exact   — height is exactly @val (overflow clipped).
//   atLeast — @val is a lower bound; content can expand the row.
//   auto    — height is content-driven; @val is IGNORED ("with no predetermined
//             minimum or maximum size"). It is only an advisory layout cache.
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
    borders: { top: null, bottom: null, left: null, right: null },
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

  it('auto — @val is IGNORED; height collapses to the content (empty ⇒ minimum)', () => {
    // The advisory cached val (600) must NOT act as a floor under auto. With no
    // content the row falls back to the 10pt minimum, not 600pt.
    const h = calculateRowHeight(rowWith(600, 'auto'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(10 * SCALE);
    expect(h).not.toBe(600 * SCALE);
  });

  it('auto with no trHeight — same as auto with a cached val (content-driven)', () => {
    const h = calculateRowHeight(rowWith(null, 'auto'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(10 * SCALE);
  });

  it('atLeast with no @val — falls back to the minimum (val null ⇒ no floor)', () => {
    const h = calculateRowHeight(rowWith(null, 'atLeast'), t, COLS, SCALE, EMPTY_STATE);
    expect(h).toBe(10 * SCALE);
  });
});
