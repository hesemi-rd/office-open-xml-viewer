/**
 * Table layout fragment production (PR 6 of the layout-context / fragments migration;
 * see docs/docx-layout-context-fragments-design.md §"Measured Fragment Model").
 *
 * `buildTableFragment` turns a (possibly page-sliced) {@link DocTable} plus its resolved
 * scale-1 geometry into an immutable {@link TableFragment}: rows fragment into
 * {@link RowFragment}s, cells into {@link CellFragment}s, and each cell's content into a
 * recursive list of {@link FlowFragment}s (paragraph fragments + nested-table fragments).
 * A fragment belongs to the {@link import('./layout-fragments.js').DocumentLayout} result,
 * never to the parsed model — the parsed {@link DocTable} is never mutated.
 *
 * This module is deliberately PURE and renderer-agnostic. The renderer-coupled work —
 * measuring a cell paragraph at scale 1, and resolving/recursing a nested table's
 * geometry — is supplied by the caller through {@link BuildCellBlocks}, so the builder
 * has no import cycle with renderer.ts and is unit-testable with a stub callback.
 *
 * Row heights and the vMerge span extension are resolved by the caller (renderer, via
 * `resolveTableRowHeights` / `resolveSingleRowHeight` / `findMergeEndRow` in
 * table-geometry.ts) and passed in as `rowHeightsPt`. They are NOT recomputed here:
 * a page slice can cut a §17.4.85 vMerge span, and the whole-table span extension grows
 * a different row than a slice-local recompute would, so only the caller — which holds
 * the whole-table context and the split — can produce heights that match the paginator's
 * cursor advancement.
 */
import type { DocTable, DocTableCell } from './types';
import type {
  CellFragment,
  FlowFragment,
  RowFragment,
  TableFragment,
} from './layout-fragments.js';

/**
 * Build the recursive content fragments of one cell laid out at `cellTotalWidthPt`
 * (the sum of the grid columns the cell spans, BEFORE the cell's own margins are
 * removed). Returns the cell's blocks in document order — a paragraph fragment per
 * `<w:p>`, a nested-table fragment per `<w:tbl>`. Supplied by the renderer, which owns
 * cell-paragraph measurement and nested-table geometry. A `vMerge=continue` cell is
 * never passed here (it renders no content); the builder gives it an empty block list.
 */
export type BuildCellBlocks = (
  cell: DocTableCell,
  cellTotalWidthPt: number,
) => readonly FlowFragment[];

export interface BuildTableFragmentInput {
  /** The (possibly page-sliced / row-split) table to fragment. Rows align 1:1 with
   *  `rowHeightsPt`. */
  readonly table: DocTable;
  /** The resolved scale-1 grid — every column of the table, constant across a
   *  page-split. */
  readonly columnWidthsPt: readonly number[];
  /** Per-row scale-1 point heights, aligned 1:1 with `table.rows` (a continuation
   *  slice's repeated-header heights are already prepended by the caller). */
  readonly rowHeightsPt: readonly number[];
  /** The source table spilled a page boundary INTO this fragment (a continuation
   *  slice). */
  readonly continuesFromPreviousPage: boolean;
  /** The source table spills a page boundary OUT of this fragment (more rows follow
   *  on a later page). */
  readonly continuesOnNextPage: boolean;
  /** Number of leading rows that are REPEATED headers on this slice (§17.4.78). 0 on
   *  the first slice / a non-continuation table. */
  readonly repeatedHeaderRowCount: number;
  /** Map a fragment row index to its index in the ORIGINAL table (a continuation
   *  slice prepends header rows; a row-split reuses one source index for several
   *  slice rows). Defaults to identity. */
  readonly sourceRowIndexOf?: (fragmentRowIndex: number) => number;
  /** Renderer-supplied cell content builder (see {@link BuildCellBlocks}). */
  readonly buildCellBlocks: BuildCellBlocks;
}

/** The §17.4.85 `<w:vMerge>` role of a cell (its parsed `vMerge`: `true`=restart,
 *  `false`=continue, `null`=none). A continue cell renders no content. */
function verticalMergeRole(cell: DocTableCell): CellFragment['verticalMerge'] {
  if (cell.vMerge === true) return 'restart';
  if (cell.vMerge === false) return 'continue';
  return 'none';
}

/**
 * Produce the immutable {@link TableFragment} for one placed table (whole table or one
 * page slice). STRUCTURALLY freezes the fragment and its nested row/cell/column arrays
 * (M-3): every wrapper object and array this module creates is frozen, so the layout
 * result's shape cannot be mutated. The freeze is structural, not deep — the referenced
 * parsed model objects and the measured paragraph internals (`MeasuredParagraph.lines`
 * and its line objects) stay shared and unfrozen, exactly as the PR 5 body fragments
 * document.
 */
export function buildTableFragment(input: BuildTableFragmentInput): TableFragment {
  const {
    table,
    columnWidthsPt,
    rowHeightsPt,
    continuesFromPreviousPage,
    continuesOnNextPage,
    repeatedHeaderRowCount,
    buildCellBlocks,
  } = input;
  const sourceRowIndexOf = input.sourceRowIndexOf ?? ((i: number) => i);
  const columns = Object.freeze([...columnWidthsPt]) as readonly number[];

  const rows: RowFragment[] = table.rows.map((sourceRow, ri) => {
    let ci = 0;
    const cells: CellFragment[] = sourceRow.cells.map((sourceCell) => {
      const span = Math.min(sourceCell.colSpan, columns.length - ci);
      let cellTotalWidthPt = 0;
      for (let cj = ci; cj < ci + span; cj++) cellTotalWidthPt += columns[cj] ?? 0;
      ci += span;
      const role = verticalMergeRole(sourceCell);
      // A continue cell renders no content (§17.4.85); a restart / none cell builds
      // its recursive blocks at the spanned width.
      const blocks =
        role === 'continue'
          ? (Object.freeze([]) as readonly FlowFragment[])
          : (Object.freeze([...buildCellBlocks(sourceCell, cellTotalWidthPt)]) as readonly FlowFragment[]);
      return Object.freeze({
        source: sourceCell,
        blocks,
        verticalMerge: role,
      }) as CellFragment;
    });
    return Object.freeze({
      source: sourceRow,
      sourceRowIndex: sourceRowIndexOf(ri),
      heightPt: rowHeightsPt[ri] ?? 0,
      cells: Object.freeze(cells) as readonly CellFragment[],
      repeatedHeader: ri < repeatedHeaderRowCount,
    }) as RowFragment;
  });

  return Object.freeze({
    kind: 'table',
    source: table,
    columnWidthsPt: columns,
    rows: Object.freeze(rows) as readonly RowFragment[],
    continuesFromPreviousPage,
    continuesOnNextPage,
  }) as TableFragment;
}
