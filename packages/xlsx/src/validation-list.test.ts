import { describe, it, expect } from 'vitest';
import {
  parseListFormula,
  resolveListValues,
  computeValidationPanelPosition,
  VALIDATION_PANEL_GAP,
} from './validation-list.js';

/**
 * `parseListFormula` classifies a list data-validation `formula1`
 * (ECMA-376 §18.3.1.32) into the three operand shapes the viewer can act on:
 * an inline quoted comma list, a cell-range reference (optionally on another
 * sheet), or something we cannot resolve (named ranges / complex formulae).
 */
describe('parseListFormula', () => {
  it('splits an inline quoted comma list', () => {
    const r = parseListFormula('"Low,Medium,High"');
    expect(r).toEqual({ kind: 'inline', values: ['Low', 'Medium', 'High'] });
  });

  it('trims whitespace around inline items but keeps interior spaces', () => {
    const r = parseListFormula('"Red , Green Apple ,Blue"');
    expect(r).toEqual({
      kind: 'inline',
      values: ['Red', 'Green Apple', 'Blue'],
    });
  });

  it('drops empty inline items from trailing / doubled commas', () => {
    const r = parseListFormula('"A,,B,"');
    expect(r).toEqual({ kind: 'inline', values: ['A', 'B'] });
  });

  it('parses a same-sheet range reference', () => {
    const r = parseListFormula('$B$2:$B$5');
    expect(r).toEqual({
      kind: 'range',
      sheet: undefined,
      start: { row: 2, col: 2 },
      end: { row: 5, col: 2 },
    });
  });

  it('parses a single-cell (degenerate) range', () => {
    const r = parseListFormula('Sheet1!$A$1');
    expect(r).toEqual({
      kind: 'range',
      sheet: 'Sheet1',
      start: { row: 1, col: 1 },
      end: { row: 1, col: 1 },
    });
  });

  it('parses a cross-sheet range with a quoted sheet name (spaces)', () => {
    const r = parseListFormula("'My Sheet'!$A$1:$A$3");
    expect(r).toEqual({
      kind: 'range',
      sheet: 'My Sheet',
      start: { row: 1, col: 1 },
      end: { row: 3, col: 1 },
    });
  });

  it('normalizes a range written from any corner', () => {
    const r = parseListFormula('$C$5:$A$2');
    expect(r).toEqual({
      kind: 'range',
      sheet: undefined,
      start: { row: 2, col: 1 },
      end: { row: 5, col: 3 },
    });
  });

  it('treats a defined name / complex formula as unresolved', () => {
    expect(parseListFormula('MyNamedList')).toEqual({
      kind: 'unresolved',
      formula: 'MyNamedList',
    });
    expect(parseListFormula('INDIRECT($A$1)')).toEqual({
      kind: 'unresolved',
      formula: 'INDIRECT($A$1)',
    });
  });

  it('treats empty / missing formula as unresolved', () => {
    expect(parseListFormula('')).toEqual({ kind: 'unresolved', formula: '' });
    expect(parseListFormula(undefined)).toEqual({
      kind: 'unresolved',
      formula: '',
    });
  });
});

/**
 * `resolveListValues` expands a parsed operand to display strings using a
 * cell-lookup callback (which the workbook backs with the real formatCellValue
 * path). Inline lists pass through; ranges are walked row-major with empties
 * skipped; unresolved operands surface the raw formula.
 */
describe('resolveListValues', () => {
  it('passes inline values straight through', () => {
    const r = resolveListValues(
      { kind: 'inline', values: ['A', 'B'] },
      () => null,
    );
    expect(r).toEqual({ kind: 'values', values: ['A', 'B'] });
  });

  it('walks a range row-major and skips empty cells', () => {
    // 2x2 block B2:C3 with C2 empty.
    const grid: Record<string, string | null> = {
      '2:2': 'Apple',
      '2:3': null,
      '3:2': 'Banana',
      '3:3': 'Cherry',
    };
    const r = resolveListValues(
      {
        kind: 'range',
        sheet: undefined,
        start: { row: 2, col: 2 },
        end: { row: 3, col: 3 },
      },
      (row, col) => grid[`${row}:${col}`] ?? null,
    );
    expect(r).toEqual({ kind: 'values', values: ['Apple', 'Banana', 'Cherry'] });
  });

  it('surfaces the raw formula for an unresolved operand', () => {
    const r = resolveListValues(
      { kind: 'unresolved', formula: 'MyNamedList' },
      () => null,
    );
    expect(r).toEqual({ kind: 'formula', formula: 'MyNamedList' });
  });
});

/**
 * The panel opens just below the dropdown arrow / cell, aligned to the cell's
 * left edge, and flips above the cell when it would overflow the bottom. RTL
 * cell rects are already mirrored by the caller (screenX), so the panel
 * right-aligns to the cell so it grows into the (mirrored) sheet body. All
 * coordinates are canvasArea CSS-pixel space.
 */
describe('computeValidationPanelPosition', () => {
  const GAP = VALIDATION_PANEL_GAP;
  const cell = { x: 100, y: 60, w: 64, h: 20 };
  const panel = { w: 120, h: 90 };
  const viewport = { w: 800, h: 600 };

  it('opens below the cell, left-aligned (LTR)', () => {
    const pos = computeValidationPanelPosition({
      cell,
      panel,
      viewport,
      rtl: false,
    });
    expect(pos.left).toBe(cell.x);
    expect(pos.top).toBe(cell.y + cell.h + GAP);
  });

  it('flips above the cell when the panel overflows the bottom (LTR)', () => {
    const bottomCell = { x: 100, y: 540, w: 64, h: 20 };
    const pos = computeValidationPanelPosition({
      cell: bottomCell,
      panel,
      viewport,
      rtl: false,
    });
    // Above: panel bottom sits GAP above the cell top.
    expect(pos.top).toBe(bottomCell.y - GAP - panel.h);
    expect(pos.top).toBeGreaterThanOrEqual(0);
  });

  it('clamps the left edge into the viewport (wide panel near right edge)', () => {
    const rightCell = { x: 760, y: 60, w: 64, h: 20 };
    const pos = computeValidationPanelPosition({
      cell: rightCell,
      panel,
      viewport,
      rtl: false,
    });
    expect(pos.left).toBeGreaterThanOrEqual(0);
    expect(pos.left + panel.w).toBeLessThanOrEqual(viewport.w);
  });

  it('right-aligns the panel to the cell for RTL', () => {
    const rtlCell = { x: 500, y: 60, w: 64, h: 20 };
    const pos = computeValidationPanelPosition({
      cell: rtlCell,
      panel,
      viewport,
      rtl: true,
    });
    // Panel right edge aligns with the cell's right edge.
    expect(pos.left).toBe(rtlCell.x + rtlCell.w - panel.w);
    expect(pos.top).toBe(rtlCell.y + rtlCell.h + GAP);
  });

  it('clamps RTL left into the viewport when the cell is near the left edge', () => {
    const leftCell = { x: 10, y: 60, w: 64, h: 20 };
    const pos = computeValidationPanelPosition({
      cell: leftCell,
      panel,
      viewport,
      rtl: true,
    });
    expect(pos.left).toBeGreaterThanOrEqual(0);
  });
});
