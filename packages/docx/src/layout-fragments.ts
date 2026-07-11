/**
 * Measured layout fragments for the DOCX body flow (PR 5 of the layout-context /
 * fragments migration; see docs/docx-layout-context-fragments-design.md §"Measured
 * Fragment Model").
 *
 * A fragment belongs to a {@link DocumentLayout} result, NOT to the parsed document
 * model — the parsed {@link DocParagraph} is never mutated with layout state. A
 * {@link ParagraphFragment} references its source paragraph plus the placement-aware
 * {@link MeasuredParagraph} that was produced for its recorded placement, and a
 * `[lineStart, lineEnd)` half-open range selecting the lines this fragment paints
 * (a paragraph split across pages/columns is represented by several fragments over
 * one measured result, each with its own range).
 *
 * All coordinates are points at scale 1 — the measurement coordinate system. Paint
 * scales the stored geometry; it never repeats text layout.
 *
 * PR 6 widens {@link FlowFragment} to also include a {@link TableFragment}: a table
 * fragments recursively into {@link RowFragment}s → {@link CellFragment}s → nested
 * {@link FlowFragment}s (paragraph or nested-table fragments), so body and table flow
 * share one immutable fragment result (design §"Measured Fragment Model").
 */
import type { DocParagraph, DocTable, DocTableRow, DocTableCell, SectionGeom } from './types';
import type { MeasuredParagraph } from './paragraph-measure.js';
import type { SectionLayoutContext } from './layout-context.js';

export interface ParagraphFragment {
  readonly kind: 'paragraph';
  /** The parsed source paragraph. Immutable — never stamped with layout state. */
  readonly source: DocParagraph;
  /** The placement-aware measurement whose lines this fragment paints. Valid only
   *  for `measured.placement`; a fragment on a different page/column/wrap context
   *  holds a separate measurement. */
  readonly measured: MeasuredParagraph;
  /** Half-open range `[lineStart, lineEnd)` into `measured.lines` this fragment
   *  paints. A markOnly measurement (empty / anchor-only paragraph) uses
   *  `lineStart === lineEnd === 0` and paints its single paragraph-mark line box. */
  readonly lineStart: number;
  readonly lineEnd: number;
  /** Points of paragraph spacing the paginator added ABOVE this fragment's first
   *  painted line (effective space-before + any float top-skip on the first
   *  fragment; 0 on a continuation fragment). Owned by the paginator so paragraph
   *  spacing is counted exactly once. */
  readonly leadingSpacePt: number;
  /** Points of paragraph spacing the paginator reserved BELOW this fragment's last
   *  painted line (max of space-after and drawn bottom-border extent on the final
   *  fragment; 0 on a non-final fragment). */
  readonly trailingSpacePt: number;
}

/**
 * A measured table cell (ECMA-376 §17.4.7 `<w:tc>`). Its content is a recursive list
 * of {@link FlowFragment}s — paragraph fragments and nested-table fragments in document
 * order — so a cell's paragraphs are painted from their own stored line partition
 * exactly like body paragraphs, and a nested table fragments recursively.
 *
 * `verticalMerge` mirrors the parsed cell's §17.4.85 `<w:vMerge>` role:
 *   - `restart` — the cell owns the merged span's content; the span's geometry is
 *     distributed across its rows by the row-height resolver.
 *   - `continue` — the cell renders no content (its `blocks` is empty); it exists so
 *     the grid/border geometry is complete.
 *   - `none` — an ordinary single-row cell.
 */
export interface CellFragment {
  readonly source: DocTableCell;
  readonly blocks: readonly FlowFragment[];
  readonly verticalMerge: 'none' | 'restart' | 'continue';
  /** Scale-1 point height of this cell's page-local row piece. This is the same
   *  post-repair row height the paginator charged and {@link RowFragment.heightPt}
   *  records; fragment-backed vAlign paint centres against it without deriving
   *  geometry again from the unsplit source cell. */
  readonly boxHeightPt?: number;
}

/**
 * A measured table row placed on one page (ECMA-376 §17.4.6 `<w:tr>`).
 *
 * `sourceRowIndex` is the row's index in the ORIGINAL {@link DocTable.rows}; `source`
 * is the (possibly cell-content-sliced) row this fragment renders — a table split
 * across pages by cell-block boundaries already produces a per-page slice of a source
 * row (renderer `splitRowByCellBlocks`). The model deliberately does NOT assume a
 * source row maps to exactly one `RowFragment`: a future mid-row / mid-cell page split
 * (§17.4.6 `cantSplit` default = splittable) may emit several `RowFragment`s that share
 * one `sourceRowIndex`, each carrying its own vertical portion via `heightPt` and its
 * cells' partial `blocks`. This PR does not implement that split; it only keeps the
 * contract open to it.
 *
 * `repeatedHeader` marks a leading `<w:tblHeader>` row re-emitted at the top of a
 * continuation page (§17.4.78). `heightPt` is the scale-1 point height the paginator
 * charged for this row.
 */
export interface RowFragment {
  readonly source: DocTableRow;
  readonly sourceRowIndex: number;
  readonly heightPt: number;
  readonly cells: readonly CellFragment[];
  readonly repeatedHeader: boolean;
}

/**
 * A measured table (ECMA-376 §17.4.4 `<w:tbl>`) placed on one page. `columnWidthsPt`
 * is the resolved scale-1 grid (constant across a page-split); `rows` are the row
 * fragments on THIS page. `continuesFromPreviousPage` / `continuesOnNextPage` record
 * whether the source table spilled a page boundary into / out of this fragment, so a
 * consumer can render header repetition and continuation without re-deriving the split.
 */
export interface TableFragment {
  readonly kind: 'table';
  readonly source: DocTable;
  readonly columnWidthsPt: readonly number[];
  readonly rows: readonly RowFragment[];
  readonly continuesFromPreviousPage: boolean;
  readonly continuesOnNextPage: boolean;
}

export type FlowFragment = ParagraphFragment | TableFragment;

export interface PlacedFragment {
  readonly fragment: FlowFragment;
  /** ECMA-376 §17.6.4 newspaper column (0-based). */
  readonly columnIndex: number;
  /** Page-absolute point coordinates (scale 1). `yPt` is the top of the fragment's
   *  leading spacing; `heightPt` is the cursor advancement the paginator charged for
   *  this fragment (== leadingSpacePt + line advances + trailingSpacePt). */
  readonly xPt: number;
  readonly yPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

export interface LayoutPage {
  readonly pageIndex: number;
  readonly section: SectionLayoutContext;
  readonly geometry: SectionGeom;
  readonly fragments: readonly PlacedFragment[];
}

export interface DocumentLayout {
  readonly pages: readonly LayoutPage[];
}

/**
 * Sum of the measured line advances this fragment paints, in scale-1 points.
 *
 * For a content fragment this is `measured.lines[i].advancePt` for the first line
 * in range plus, for each subsequent line, the (non-negative) inter-line gap to the
 * previous line's bottom followed by its advance — the same per-line extent the
 * paginator charges. For a markOnly fragment it is the paragraph-mark line box
 * height. It excludes the fragment's leading and trailing paragraph spacing (those
 * are `leadingSpacePt` / `trailingSpacePt`), so spacing is never double-counted.
 */
export function fragmentLineAdvancesPt(fragment: ParagraphFragment): number {
  const measured = fragment.measured;
  if (measured.markOnly || measured.lines.length === 0) {
    return measured.contentEndYPt - measured.contentStartYPt;
  }
  let sum = 0;
  for (let i = fragment.lineStart; i < fragment.lineEnd; i++) {
    const line = measured.lines[i];
    if (!line) break;
    if (i === fragment.lineStart) {
      sum += line.advancePt;
      continue;
    }
    const previous = measured.lines[i - 1];
    const previousBottom = previous.topYPt + previous.advancePt;
    sum += Math.max(0, line.topYPt - previousBottom) + line.advancePt;
  }
  return sum;
}

/**
 * The cursor advancement a paragraph fragment charges: its leading spacing, plus the
 * measured line advances it paints, plus its trailing spacing. This is the invariant
 * a paginator's per-fragment height must equal — proving paragraph spacing is owned
 * by the fragment and added exactly once (design §"Pagination and paint invariants" 1).
 */
export function paragraphFragmentAdvancePt(fragment: ParagraphFragment): number {
  return (
    fragment.leadingSpacePt +
    fragmentLineAdvancesPt(fragment) +
    fragment.trailingSpacePt
  );
}

/**
 * The flow height (scale-1 pt) a table fragment charges: the sum of its row-fragment
 * heights. Each {@link RowFragment.heightPt} is the ST_HeightRule + §17.4.85 vMerge
 * span height the paginator resolved, so this equals the cursor advancement the
 * paginator charged for the placed slice (design §"Pagination and paint invariants" 1).
 */
export function tableFragmentHeightPt(fragment: TableFragment): number {
  let sum = 0;
  for (const row of fragment.rows) sum += row.heightPt;
  return sum;
}

/** Piece-local content height (scale-1 pt) owned by a cell fragment. Paragraph
 * blocks contribute only their recorded `[lineStart, lineEnd)` advancement plus
 * fragment-owned spacing; nested tables contribute their already-resolved row
 * heights. This is deliberately a pure fragment fold: paint must not call a cell
 * or paragraph measurer to recover geometry the paginator already produced. */
export function cellFragmentContentHeightPt(fragment: CellFragment): number {
  let sum = 0;
  for (const block of fragment.blocks) {
    sum += block.kind === 'table'
      ? tableFragmentHeightPt(block)
      : paragraphFragmentAdvancePt(block);
  }
  return sum;
}

/** Flow height (scale-1 pt) of any {@link FlowFragment}: a paragraph fragment's
 *  leading + line advances + trailing, or a table fragment's summed row heights. */
export function flowFragmentAdvancePt(fragment: FlowFragment): number {
  return fragment.kind === 'table'
    ? tableFragmentHeightPt(fragment)
    : paragraphFragmentAdvancePt(fragment);
}
