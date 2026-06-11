import type { DataValidation } from './types.js';
import { parseA1 } from './a1.js';

/**
 * Test whether a 1-based (row, col) cell falls inside a `@sqref` (ECMA-376
 * §18.3.1.33). `sqref` is a space-separated list of A1 ranges ("A1",
 * "C3:E6"); ranges may be written from any corner (Excel emits them anchored
 * at the active cell, so the start corner can be below/right of the end).
 */
export function cellInSqref(sqref: string, row: number, col: number): boolean {
  if (!sqref) return false;
  for (const token of sqref.split(/\s+/)) {
    if (!token) continue;
    const [a, b] = token.split(':');
    const start = parseA1(a);
    if (!start) continue;
    if (!b) {
      if (start.row === row && start.col === col) return true;
      continue;
    }
    const end = parseA1(b);
    if (!end) continue;
    const r1 = Math.min(start.row, end.row);
    const r2 = Math.max(start.row, end.row);
    const c1 = Math.min(start.col, end.col);
    const c2 = Math.max(start.col, end.col);
    if (row >= r1 && row <= r2 && col >= c1 && col <= c2) return true;
  }
  return false;
}

/**
 * Find the first `list`-type data-validation rule whose `@sqref` covers the
 * given 1-based cell, or null. Only `list` rules drive the dropdown arrow —
 * Excel shows the in-cell dropdown button exclusively for list validation
 * (ECMA-376 §18.3.1.33, `ST_DataValidationType` value `list`); whole / decimal
 * / date / textLength / custom rules have no dropdown.
 */
export function findListValidationAt(
  validations: DataValidation[] | undefined,
  row: number,
  col: number,
): DataValidation | null {
  if (!validations) return null;
  for (const dv of validations) {
    if (dv.validationType !== 'list') continue;
    if (cellInSqref(dv.sqref, row, col)) return dv;
  }
  return null;
}
