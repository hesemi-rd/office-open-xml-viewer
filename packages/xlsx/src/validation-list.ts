import { parseA1 } from './a1.js';

/**
 * Pure helpers for the `list`-type data-validation dropdown (ECMA-376
 * §18.3.1.32). Kept DOM-free so the operand parsing and panel geometry can be
 * unit-tested; the viewer owns the actual overlay element and the workbook owns
 * the cell-value lookup for range references.
 */

/** Gap (CSS px) between the cell edge and the dropdown panel. */
export const VALIDATION_PANEL_GAP = 2;

/** A1 cell coordinate (1-based). */
interface RC {
  row: number;
  col: number;
}

/**
 * Classification of a list `formula1` operand:
 * - `inline`  — a quoted literal list `"A,B,C"`; values are split out.
 * - `range`   — a cell-range reference, optionally qualified by a sheet name;
 *               corners are normalized so `start` is the top-left.
 * - `unresolved` — a defined name, INDIRECT(), or any form we cannot expand to
 *               concrete cells; the raw formula text is carried through so the
 *               panel can show it rather than silently blanking.
 */
export type ListFormula =
  | { kind: 'inline'; values: string[] }
  | { kind: 'range'; sheet: string | undefined; start: RC; end: RC }
  | { kind: 'unresolved'; formula: string };

/**
 * Parse a list-validation `formula1` into a {@link ListFormula}.
 *
 * Excel stores the literal list quoted (`"Low,Medium,High"`); a reference is an
 * A1 range that may carry a sheet qualifier (`Sheet2!$A$1:$A$9`, or quoted when
 * the name has spaces: `'My Sheet'!$A$1:$A$3`). Everything else (named ranges,
 * INDIRECT, etc.) we leave unresolved.
 */
export function parseListFormula(formula1: string | undefined): ListFormula {
  const raw = (formula1 ?? '').trim();
  if (!raw) return { kind: 'unresolved', formula: '' };

  // Inline quoted list: "A,B,C". Excel always double-quotes the literal form.
  if (raw.length >= 2 && raw.startsWith('"') && raw.endsWith('"')) {
    const inner = raw.slice(1, -1);
    const values = inner
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s.length > 0);
    return { kind: 'inline', values };
  }

  // Optional sheet qualifier: `Sheet2!...` or `'My Sheet'!...`.
  let sheet: string | undefined;
  let rest = raw;
  const bang = raw.indexOf('!');
  if (bang >= 0) {
    let name = raw.slice(0, bang);
    if (name.startsWith("'") && name.endsWith("'") && name.length >= 2) {
      // Unescape Excel's doubled single-quote inside a quoted sheet name.
      name = name.slice(1, -1).replace(/''/g, "'");
    }
    sheet = name;
    rest = raw.slice(bang + 1);
  }

  // A1 range (possibly a single cell). parseA1 strips `$` markers.
  const [aRef, bRef] = rest.split(':');
  const a = parseA1(aRef ?? '');
  if (a) {
    const b = bRef ? parseA1(bRef) : a;
    if (b) {
      const start: RC = {
        row: Math.min(a.row, b.row),
        col: Math.min(a.col, b.col),
      };
      const end: RC = {
        row: Math.max(a.row, b.row),
        col: Math.max(a.col, b.col),
      };
      return { kind: 'range', sheet, start, end };
    }
  }

  return { kind: 'unresolved', formula: raw };
}

/**
 * Resolved allowed-value set for a list validation. Either concrete display
 * `values` (inline list or expanded range), or — when the operand is a defined
 * name / complex formula we cannot expand — the raw `formula` text so the panel
 * can disclose it instead of showing nothing.
 */
export type ResolvedList =
  | { kind: 'values'; values: string[] }
  | { kind: 'formula'; formula: string };

/**
 * Expand a parsed list operand to its display values, given a callback that
 * returns the formatted display text of a cell (or null for empty / missing).
 * Pure: the workbook supplies a `cellAt` closure that reads the right sheet and
 * runs the existing `formatCellValue` path, but the iteration / ordering /
 * empty-skipping lives here so it can be unit-tested.
 *
 * Range cells are walked row-major (top→bottom, left→right) to match the order
 * Excel lists them in the dropdown. Empty cells are skipped (Excel omits blanks
 * from the in-cell list).
 */
export function resolveListValues(
  parsed: ListFormula,
  cellAt: (row: number, col: number) => string | null,
): ResolvedList {
  if (parsed.kind === 'inline') {
    return { kind: 'values', values: parsed.values };
  }
  if (parsed.kind === 'unresolved') {
    return { kind: 'formula', formula: parsed.formula };
  }
  const values: string[] = [];
  for (let r = parsed.start.row; r <= parsed.end.row; r++) {
    for (let c = parsed.start.col; c <= parsed.end.col; c++) {
      const text = cellAt(r, c);
      if (text != null && text !== '') values.push(text);
    }
  }
  return { kind: 'values', values };
}

export interface ValidationPanelGeometry {
  /** The validated cell's on-screen rect (canvasArea space, RTL already mirrored). */
  cell: { x: number; y: number; w: number; h: number };
  /** The panel's measured size. */
  panel: { w: number; h: number };
  /** The visible viewport (canvasArea client size). */
  viewport: { w: number; h: number };
  /** True when the current sheet is laid out right-to-left. */
  rtl: boolean;
}

/**
 * Compute the top-left of the dropdown panel, in canvasArea CSS pixels.
 *
 * Vertical: open below the cell (anchored at the cell's bottom edge + gap). If
 * the panel would overflow the bottom, flip above the cell. Then clamp to stay
 * fully inside [0, viewport.h].
 *
 * Horizontal: LTR left-aligns to the cell's left edge; RTL right-aligns to the
 * cell's right edge (its rect is already mirrored, so this grows the panel into
 * the sheet body). Finally clamp into [0, viewport.w].
 */
export function computeValidationPanelPosition(geo: ValidationPanelGeometry): {
  left: number;
  top: number;
} {
  const { cell, panel, viewport, rtl } = geo;
  const gap = VALIDATION_PANEL_GAP;

  const below = cell.y + cell.h + gap;
  const above = cell.y - gap - panel.h;
  let top: number;
  if (below + panel.h <= viewport.h) top = below;
  else if (above >= 0) top = above;
  else top = below;
  top = Math.max(0, Math.min(top, viewport.h - panel.h));

  let left = rtl ? cell.x + cell.w - panel.w : cell.x;
  left = Math.max(0, Math.min(left, viewport.w - panel.w));

  return { left, top };
}
