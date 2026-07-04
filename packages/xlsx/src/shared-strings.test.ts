import { describe, it, expect } from 'vitest';
import { resolveSharedStrings } from './shared-strings.js';
import type { SharedString, Worksheet, Cell, Run } from './types.js';

/** Build a minimal single-row Worksheet carrying the given cells. Only the
 *  fields `resolveSharedStrings` touches (`rows[].cells[].value`) matter; the
 *  rest are filled with harmless defaults so the object type-checks. */
function ws(cells: Cell[]): Worksheet {
  return {
    name: 'Sheet1',
    rows: [{ index: 1, height: null, cells }],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  };
}

function sharedCell(col: number, si: number): Cell {
  return { col, row: 1, value: { type: 'shared', si } };
}

describe('resolveSharedStrings', () => {
  it('resolves si=0 (boundary) to the first table entry', () => {
    const table: SharedString[] = [{ text: 'first' }, { text: 'second' }];
    const sheet = resolveSharedStrings(ws([sharedCell(1, 0)]), table);
    expect(sheet.rows[0].cells[0].value).toEqual({ type: 'text', text: 'first' });
  });

  it('resolves a mid-table index', () => {
    const table: SharedString[] = [{ text: 'a' }, { text: 'b' }, { text: 'c' }];
    const sheet = resolveSharedStrings(ws([sharedCell(1, 2)]), table);
    expect(sheet.rows[0].cells[0].value).toEqual({ type: 'text', text: 'c' });
  });

  it('resolves an out-of-range si to empty text (parser fallback parity)', () => {
    const table: SharedString[] = [{ text: 'only' }];
    const sheet = resolveSharedStrings(ws([sharedCell(1, 9)]), table);
    // Matches the historical Rust fallback: a missing index → empty string,
    // no runs.
    expect(sheet.rows[0].cells[0].value).toEqual({ type: 'text', text: '' });
  });

  it('resolves against an empty table to empty text', () => {
    const sheet = resolveSharedStrings(ws([sharedCell(1, 0)]), []);
    expect(sheet.rows[0].cells[0].value).toEqual({ type: 'text', text: '' });
  });

  it('resolves an empty-string shared entry to empty text (no runs key)', () => {
    const table: SharedString[] = [{ text: '' }];
    const sheet = resolveSharedStrings(ws([sharedCell(1, 0)]), table);
    const v = sheet.rows[0].cells[0].value;
    expect(v).toEqual({ type: 'text', text: '' });
    // A shared string with no runs must NOT introduce a `runs` key — the
    // resolved shape must match exactly what the parser used to emit inline.
    expect('runs' in (v as { runs?: Run[] })).toBe(false);
  });

  it('preserves multiple rich-text runs verbatim', () => {
    const runs: Run[] = [
      { text: 'Bold', font: { bold: true, italic: false, underline: false, strike: false } },
      { text: ' plain' },
      { text: ' red', font: { bold: false, italic: false, underline: false, strike: false, color: '#FF0000' } },
    ];
    const table: SharedString[] = [{ text: 'Bold plain red', runs }];
    const sheet = resolveSharedStrings(ws([sharedCell(1, 0)]), table);
    expect(sheet.rows[0].cells[0].value).toEqual({
      type: 'text',
      text: 'Bold plain red',
      runs,
    });
    // Same array reference is fine (the table lives for the workbook lifetime),
    // but the run contents must round-trip identically.
    expect((sheet.rows[0].cells[0].value as { runs?: Run[] }).runs).toEqual(runs);
  });

  it('leaves inline / non-shared cells untouched while resolving shared ones', () => {
    const table: SharedString[] = [{ text: 'SHARED' }];
    const inline: Cell = { col: 2, row: 1, value: { type: 'text', text: 'inline', runs: [{ text: 'inline' }] } };
    const num: Cell = { col: 3, row: 1, value: { type: 'number', number: 42 } };
    const empty: Cell = { col: 4, row: 1, value: { type: 'empty' } };
    const sheet = resolveSharedStrings(ws([sharedCell(1, 0), inline, num, empty]), table);
    expect(sheet.rows[0].cells[0].value).toEqual({ type: 'text', text: 'SHARED' });
    // Non-shared values pass through by identity.
    expect(sheet.rows[0].cells[1].value).toEqual({ type: 'text', text: 'inline', runs: [{ text: 'inline' }] });
    expect(sheet.rows[0].cells[2].value).toEqual({ type: 'number', number: 42 });
    expect(sheet.rows[0].cells[3].value).toEqual({ type: 'empty' });
  });

  it('is idempotent: a worksheet with no shared cells is returned unchanged', () => {
    const num: Cell = { col: 1, row: 1, value: { type: 'number', number: 7 } };
    const sheet = ws([num]);
    const before = JSON.stringify(sheet);
    const out = resolveSharedStrings(sheet, [{ text: 'x' }]);
    expect(out).toBe(sheet); // same object (mutated in place, returned)
    expect(JSON.stringify(out)).toBe(before);
  });
});
