import type {
  Workbook,
  Worksheet,
  Cell,
  CellValue,
  Row,
} from '@silurus/ooxml-xlsx';
import type { Change, DiffResult } from './types.ts';
import { deepEqual } from './util/equal.ts';

/** Input to {@link diffXlsx}: the workbook plus a map of sheet name → parsed Worksheet.
 *  This mirrors how `XlsxWorkbook` exposes worksheets (one parse per sheet on demand). */
export interface XlsxDiffInput {
  workbook: Workbook;
  worksheets: Record<string, Worksheet>;
}

/** Top-level entry. Compare two XLSX workbooks. */
export function diffXlsx(before: XlsxDiffInput, after: XlsxDiffInput): DiffResult {
  const changes: Change[] = [];

  diffSheetList(before.workbook, after.workbook, changes);

  const sheetNames = new Set<string>();
  for (const s of before.workbook.sheets) sheetNames.add(s.name);
  for (const s of after.workbook.sheets) sheetNames.add(s.name);

  for (const name of sheetNames) {
    const a = before.worksheets[name];
    const b = after.worksheets[name];
    if (a == null && b != null) {
      changes.push({
        op: 'add',
        path: `sheets["${name}"]`,
        kind: 'sheet',
        after: b,
        location: { kind: 'sheet', sheetName: name },
      });
      continue;
    }
    if (a != null && b == null) {
      changes.push({
        op: 'remove',
        path: `sheets["${name}"]`,
        kind: 'sheet',
        before: a,
        location: { kind: 'sheet', sheetName: name },
      });
      continue;
    }
    if (a == null || b == null) continue;
    diffWorksheet(a, b, name, changes);
  }

  return { format: 'xlsx', changes };
}

function diffSheetList(a: Workbook, b: Workbook, out: Change[]): void {
  const aNames = a.sheets.map((s) => s.name);
  const bNames = b.sheets.map((s) => s.name);
  if (!deepEqual(aNames, bNames)) {
    out.push({
      op: 'modify',
      path: 'sheets',
      kind: 'sheet-list',
      before: aNames,
      after: bNames,
    });
  }
}

function diffWorksheet(a: Worksheet, b: Worksheet, sheetName: string, out: Change[]): void {
  diffCells(a.rows, b.rows, sheetName, out);

  // Merge cells
  if (!deepEqual(a.mergeCells, b.mergeCells)) {
    out.push({
      op: 'modify',
      path: `sheets["${sheetName}"].mergeCells`,
      kind: 'merges',
      before: a.mergeCells,
      after: b.mergeCells,
      location: { kind: 'sheet', sheetName },
    });
  }

  // Column widths
  if (!deepEqual(a.colWidths, b.colWidths)) {
    out.push({
      op: 'modify',
      path: `sheets["${sheetName}"].colWidths`,
      kind: 'col-widths',
      before: a.colWidths,
      after: b.colWidths,
      location: { kind: 'sheet', sheetName },
    });
  }

  // Row heights
  if (!deepEqual(a.rowHeights, b.rowHeights)) {
    out.push({
      op: 'modify',
      path: `sheets["${sheetName}"].rowHeights`,
      kind: 'row-heights',
      before: a.rowHeights,
      after: b.rowHeights,
      location: { kind: 'sheet', sheetName },
    });
  }

  // Freeze pane
  if (a.freezeRows !== b.freezeRows || a.freezeCols !== b.freezeCols) {
    out.push({
      op: 'modify',
      path: `sheets["${sheetName}"].freeze`,
      kind: 'freeze',
      before: { rows: a.freezeRows, cols: a.freezeCols },
      after: { rows: b.freezeRows, cols: b.freezeCols },
      location: { kind: 'sheet', sheetName },
    });
  }
}

function diffCells(beforeRows: Row[], afterRows: Row[], sheetName: string, out: Change[]): void {
  const beforeMap = buildCellMap(beforeRows);
  const afterMap = buildCellMap(afterRows);

  const keys = new Set<string>([...beforeMap.keys(), ...afterMap.keys()]);
  for (const key of keys) {
    const a = beforeMap.get(key);
    const b = afterMap.get(key);
    if (a == null && b != null) {
      out.push({
        op: 'add',
        path: `sheets["${sheetName}"].cells.${key}`,
        kind: 'cell',
        after: cellSummary(b),
        location: { kind: 'cell', sheetName, row: b.row, col: b.col },
      });
      continue;
    }
    if (a != null && b == null) {
      out.push({
        op: 'remove',
        path: `sheets["${sheetName}"].cells.${key}`,
        kind: 'cell',
        before: cellSummary(a),
        location: { kind: 'cell', sheetName, row: a.row, col: a.col },
      });
      continue;
    }
    if (a == null || b == null) continue;

    if (!cellValueEqual(a.value, b.value)) {
      out.push({
        op: 'modify',
        path: `sheets["${sheetName}"].cells.${key}.value`,
        kind: 'cell-value',
        before: cellSummary(a),
        after: cellSummary(b),
        location: { kind: 'cell', sheetName, row: b.row, col: b.col },
      });
    } else if (a.styleIndex !== b.styleIndex) {
      out.push({
        op: 'modify',
        path: `sheets["${sheetName}"].cells.${key}.styleIndex`,
        kind: 'cell-style',
        before: a.styleIndex,
        after: b.styleIndex,
        location: { kind: 'cell', sheetName, row: b.row, col: b.col },
      });
    } else if (a.formula !== b.formula) {
      out.push({
        op: 'modify',
        path: `sheets["${sheetName}"].cells.${key}.formula`,
        kind: 'cell-formula',
        before: a.formula,
        after: b.formula,
        location: { kind: 'cell', sheetName, row: b.row, col: b.col },
      });
    }
  }
}

function buildCellMap(rows: Row[]): Map<string, Cell> {
  const m = new Map<string, Cell>();
  for (const row of rows) {
    for (const c of row.cells) {
      m.set(c.colRef, c);
    }
  }
  return m;
}

function cellValueEqual(a: CellValue, b: CellValue): boolean {
  if (a.type !== b.type) return false;
  switch (a.type) {
    case 'empty':
      return true;
    case 'text':
      return a.text === (b as { text: string }).text;
    case 'number':
      return a.number === (b as { number: number }).number;
    case 'bool':
      return a.bool === (b as { bool: boolean }).bool;
    case 'error':
      return a.error === (b as { error: string }).error;
    default:
      return false;
  }
}

function cellSummary(c: Cell): string {
  switch (c.value.type) {
    case 'empty':
      return '';
    case 'text':
      return c.value.text;
    case 'number':
      return String(c.value.number);
    case 'bool':
      return String(c.value.bool);
    case 'error':
      return `#${c.value.error}`;
    default:
      return '';
  }
}
