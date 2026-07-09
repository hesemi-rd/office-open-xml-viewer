import { describe, it, expect } from 'vitest';
import { splitTableAcrossPages } from './renderer.js';
import type { DocTable, DocTableRow, PaginatedBodyElement } from './types';

// Minimal row/table builders — splitTableAcrossPages only reads rows/cells
// (vMerge, isHeader) and the precomputed rowHs, never measures.
function row(opts: { isHeader?: boolean; vMergeContinue?: boolean } = {}): DocTableRow {
  return {
    cells: [
      {
        content: [],
        colSpan: 1,
        vMerge: opts.vMergeContinue ? false : null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null,
        vAlign: 'top',
        widthPt: null,
      },
    ],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: opts.isHeader ?? false,
  };
}

function table(rows: DocTableRow[]): DocTable {
  return {
    colWidths: [100],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
}

/** Drive splitTableAcrossPages with a fresh pages array + a newPage that mirrors
 *  computePages (append an empty page). Returns the resulting pages. */
function run(t: DocTable, rowHs: number[], startY: number, contentH: number) {
  const pages: PaginatedBodyElement[][] = [[]];
  const newPage = () => { pages.push([]); };
  const endY = splitTableAcrossPages(t, rowHs, startY, contentH, pages, newPage);
  return { pages, endY };
}

const rowsOf = (slice: PaginatedBodyElement) => (slice as unknown as DocTable).rows;

describe('splitTableAcrossPages', () => {
  it('splits a tall table into per-page slices that each fit, preserving all rows in order', () => {
    const t = table(Array.from({ length: 10 }, () => row()));
    const rowHs = Array(10).fill(100);
    const { pages } = run(t, rowHs, 0, 350); // 3 rows (300) fit per 350 page

    // 10 rows → 3 + 3 + 3 + 1 across 4 pages.
    expect(pages.length).toBe(4);
    expect(pages.map((p) => rowsOf(p[0]).length)).toEqual([3, 3, 3, 1]);
    // Every page slice fits within contentH.
    for (const p of pages) {
      const h = rowsOf(p[0]).length * 100;
      expect(h).toBeLessThanOrEqual(350);
    }
    // No row lost or duplicated.
    expect(pages.reduce((s, p) => s + rowsOf(p[0]).length, 0)).toBe(10);
  });

  it('fills the remaining space on the starting page before breaking', () => {
    const t = table(Array.from({ length: 6 }, () => row()));
    const rowHs = Array(6).fill(100);
    // Start partway down a 350-tall page: only 1 row (100) fits in the 150 left.
    const { pages } = run(t, rowHs, 200, 350);
    expect(rowsOf(pages[0][0]).length).toBe(1);
    expect(pages.length).toBe(3); // 1 + 3 + 2
    expect(pages.reduce((s, p) => s + rowsOf(p[0]).length, 0)).toBe(6);
  });

  it('moves a row to the next page when it cannot fit in the remaining band', () => {
    const t = table(Array.from({ length: 4 }, () => row()));
    const rowHs = Array(4).fill(100);
    // Only 50pt remains on the current page. A 100pt row fits on a fresh page, so
    // it must not be emitted into the footer/page-number band of the current one.
    const { pages } = run(t, rowHs, 300, 350);

    expect(pages[0]).toHaveLength(0);
    expect(pages[1]).toHaveLength(1);
    expect(rowsOf(pages[1][0]).length).toBe(3);
    expect(rowsOf(pages[2][0]).length).toBe(1);
  });

  it('repeats leading tblHeader rows at the top of every continuation page', () => {
    const header = row({ isHeader: true });
    const body = Array.from({ length: 8 }, () => row());
    const t = table([header, ...body]);
    const rowHs = Array(9).fill(100);
    const { pages } = run(t, rowHs, 0, 350);

    // First page: header + first body rows. Continuations re-prepend the header.
    expect(rowsOf(pages[0][0])[0].isHeader).toBe(true);
    for (let i = 1; i < pages.length; i++) {
      expect(rowsOf(pages[i][0])[0].isHeader).toBe(true);
    }
  });

  it('never breaks before a vMerge continuation row (keeps the merged span together)', () => {
    // Rows 0..4; row 3 continues a vertical merge started at row 2.
    const rows = [row(), row(), row(), row({ vMergeContinue: true }), row()];
    const t = table(rows);
    const rowHs = Array(5).fill(100);
    // 350 page would naturally break after row 2 (300) — but the next page would
    // start at row 3, a vMerge continuation. The break must not orphan it: with
    // forward-only progress the break still lands on a safe boundary.
    const { pages } = run(t, rowHs, 0, 350);
    // Reconstruct the row order across slices (dropping repeated headers — none here).
    const flat = pages.flatMap((p) => rowsOf(p[0]));
    expect(flat.length).toBe(5);
    // For each page boundary, the first row of a continuation slice must not be
    // a vMerge continuation row (a row carrying any vMerge=false cell).
    for (let i = 1; i < pages.length; i++) {
      const firstRow = rowsOf(pages[i][0])[0];
      expect(firstRow.cells.some((c) => c.vMerge === false)).toBe(false);
    }
  });

  it('backs up to the previous safe boundary instead of overflowing before a vMerge continuation', () => {
    // Row 2 starts a vertical merge and row 3 continues it. A 350pt page can fit
    // rows 0..2 but cannot also fit row 3. ECMA-376 §17.4.85 makes a break before
    // row 3 unsafe, so the paginator must break earlier, before the merged span.
    const rows = [row(), row(), row(), row({ vMergeContinue: true }), row()];
    const t = table(rows);
    const rowHs = Array(5).fill(100);
    const { pages } = run(t, rowHs, 0, 350);

    expect(pages.map((p) => rowsOf(p[0]).length)).toEqual([2, 3]);
  });
});
