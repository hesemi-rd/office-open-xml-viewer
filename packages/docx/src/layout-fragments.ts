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
 * PR 6 will widen {@link FlowFragment} to also include a `TableFragment`. This module
 * intentionally exposes only the paragraph fragment contract for PR 5.
 */
import type { DocParagraph, SectionGeom } from './types';
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

export type FlowFragment = ParagraphFragment;

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
