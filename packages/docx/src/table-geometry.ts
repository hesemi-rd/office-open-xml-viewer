// Table row-height resolution skeleton (ECMA-376 §17.4.80 `<w:trHeight>` /
// §17.18.37 ST_HeightRule + §17.4.85 `<w:vMerge>` span extension).
//
// The trHeight rule (exact / atLeast / auto), the auto/no-floor minimum, the
// gridSpan column slicing, and the vMerge restart-span post-pass are pure
// structural math that does NOT depend on whether the caller works in pt
// (paginator) or px (paint) units. This module factors that skeleton out of
// renderer.ts so the rule is expressed once: previously both
// `computeTableRowHeights` (pt) and `computeTableLayout`/`calculateRowHeight`
// (px) re-implemented it, and #523 had to patch the identical regression in two
// places. The two callers still differ in HOW a cell's content height is
// measured (the paginator's `estimateParagraphHeight` cursor-walk vs the paint
// pass's `measureParaHeight`); that measurement is supplied as a callback and is
// deliberately NOT unified here — see the note on `measureCellContentHeight`.
//
// Only DocTable/DocTableRow/DocTableCell types are imported (erased at runtime),
// so there is no import cycle with renderer.ts.

import type { DocTable, DocTableRow, DocTableCell } from './types.js';

/** Minimum table-row height (pt) when no `w:trHeight` floor applies — i.e. an
 *  `auto` row, or `atLeast`/`exact` with no `@val`. ECMA-376 leaves the auto
 *  minimum implementation-defined; this is the floor an empty row collapses to
 *  before content (cell margins + measured content) expands it. */
export const MIN_ROW_HEIGHT_PT = 10;

/** Last row index covered by the vMerge span that starts at (`startRi`,
 *  `startCi`). A span continues through following rows whose cell anchored in the
 *  same grid column carries `vMerge=continue` (ECMA-376 §17.4.85). Pure
 *  table-structure walk — no geometry. */
export function findMergeEndRow(table: DocTable, startRi: number, startCi: number): number {
  let endRi = startRi;
  for (let rj = startRi + 1; rj < table.rows.length; rj++) {
    const row = table.rows[rj];
    let ci = 0;
    let matched = false;
    for (const cell of row.cells) {
      if (ci === startCi) {
        if (cell.vMerge === false) matched = true;
        break;
      }
      if (ci > startCi) break;
      ci += cell.colSpan;
    }
    if (!matched) break;
    endRi = rj;
  }
  return endRi;
}

/** Measure the content height (already in the target units — px at `scale`, or
 *  pt at `scale=1`) of a single cell laid out at the given total cell width
 *  (`cellWidth`, the summed widths of the grid columns it spans). MUST include
 *  the cell's top/bottom margins. The two callers pass DIFFERENT measurers — the
 *  paginator mirrors `renderParagraph`'s float-aware cursor walk
 *  (`estimateParagraphHeight`); the paint pass uses the lighter `measureParaHeight`
 *  + explicit spaceBefore/After. They are not equivalent (e.g. an empty East
 *  Asian paragraph mark on a docGrid rounds to a different cell count), so this
 *  skeleton keeps each caller's measurer intact rather than choosing one. */
export type MeasureCellContentHeight = (cell: DocTableCell, cellWidth: number) => number;

/**
 * Resolve per-row heights for a table whose grid columns have widths
 * `colWidths` (in the same target units as the measurer returns: px when
 * `scale` is the device scale, pt when `scale === 1`). Applies, once:
 *
 *   - ECMA-376 §17.4.80 / §17.18.37 (ST_HeightRule):
 *       exact   — height is exactly `w:trHeight/@val` (× `scale`); overflow is
 *                 clipped by the caller.
 *       atLeast — `@val` (× `scale`) is a lower bound; content can grow the row.
 *       auto    — `@val` is IGNORED ("no predetermined minimum or maximum size",
 *                 advisory layout cache only); the row falls back to the
 *                 `MIN_ROW_HEIGHT_PT` floor. Same as a row with no `w:trHeight`.
 *   - gridSpan: a cell's width is the sum of the `cell.colSpan` columns it
 *     anchors (clamped to the remaining columns).
 *   - ECMA-376 §17.4.85 (vMerge): a `vMerge=restart` cell's content occupies the
 *     whole merged span. It is EXCLUDED from its first row's height (so the first
 *     row is not inflated) and instead a post-pass extends the span's LAST row
 *     when the restart cell's content exceeds the summed span height.
 *     `vMerge=continue` cells render no content.
 *
 * `measureCellContentHeight` supplies the unit-specific content measurement (see
 * its type doc). The restart cell is measured through the SAME callback in the
 * post-pass; the callback must be a pure read of layout state (it is, in both
 * callers) so re-measuring yields the value the first pass would have computed.
 */
/**
 * Height of ONE row (ECMA-376 §17.4.80 / §17.18.37 ST_HeightRule + gridSpan),
 * EXCLUDING the §17.4.85 vMerge span extension (the caller / the table-level
 * resolver applies that in a post-pass). `exact` returns exactly `@val × scale`;
 * `atLeast` floors at `@val × scale`; `auto` / no-`@val` floors at
 * `MIN_ROW_HEIGHT_PT × scale`. A `vMerge=restart` cell is excluded (its content
 * is distributed across the span, not absorbed by its first row) and a
 * `vMerge=continue` cell renders no content. This is the single source of the
 * trHeight rule, shared by {@link resolveTableRowHeights} and the exported
 * `calculateRowHeight`.
 */
export function resolveSingleRowHeight(
  row: DocTableRow,
  colWidths: number[],
  scale: number,
  measureCellContentHeight: MeasureCellContentHeight,
): number {
  if (row.rowHeight != null && row.rowHeightRule === 'exact') return row.rowHeight * scale;
  let rowH =
    row.rowHeight != null && row.rowHeightRule === 'atLeast'
      ? row.rowHeight * scale
      : MIN_ROW_HEIGHT_PT * scale;

  let ci = 0;
  for (const cell of row.cells) {
    const span = Math.min(cell.colSpan, colWidths.length - ci);
    // vMerge=restart cells are sized by the span post-pass; vMerge=continue
    // cells render no content. Neither raises THIS row's height directly.
    if (cell.vMerge !== true && cell.vMerge !== false) {
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
      const ch = measureCellContentHeight(cell, cellW);
      if (ch > rowH) rowH = ch;
    }
    ci += span;
  }
  return rowH;
}

export function resolveTableRowHeights(
  table: DocTable,
  colWidths: number[],
  scale: number,
  measureCellContentHeight: MeasureCellContentHeight,
): number[] {
  const rowHeights = table.rows.map((row) =>
    resolveSingleRowHeight(row, colWidths, scale, measureCellContentHeight),
  );

  // §17.4.85 span extension: for each vMerge=restart cell, grow the span's last
  // row if the restart cell's full content is taller than the summed span rows.
  for (let ri = 0; ri < table.rows.length; ri++) {
    let ci = 0;
    for (const cell of table.rows[ri].cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      if (cell.vMerge === true) {
        const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
        const contentH = measureCellContentHeight(cell, cellW);
        const endRi = findMergeEndRow(table, ri, ci);
        let spanH = 0;
        for (let rj = ri; rj <= endRi; rj++) spanH += rowHeights[rj];
        if (spanH < contentH) {
          rowHeights[endRi] += contentH - spanH;
        }
      }
      ci += span;
    }
  }

  return rowHeights;
}
