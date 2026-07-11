/**
 * Body paragraph fragment paint (PR 5 Task 13).
 *
 * `paintParagraphFragment` draws a {@link PlacedFragment} produced by
 * {@link layoutDocument} / the paginator: it hands the fragment's stored scale-1
 * line partition to the renderer's shared paragraph draw path, which rescales it to
 * the paint scale through the existing `rescaleLayoutLines` bridge and draws only the
 * fragment's `[lineStart, lineEnd)` range. No line layout, segment construction, text
 * measurement, or paragraph re-measurement happens here — the geometry is already
 * measured. A scale-1 paint therefore invokes no `measureText` at all (the rescale
 * bridge returns the partition unchanged at scale 1).
 *
 * This module is the static body-paint boundary: the ast-grep rule
 * `no-docx-measurement-in-fragment-paint` forbids importing or calling `buildSegments`,
 * `layoutLines`, `measureParagraph`, `measureText`, or row measurement from this file.
 *
 * BOUNDARY SCOPE (PR 5): the rule enforces the measurement-free boundary for THIS
 * module only. The shared draw path this delegates to (`renderBodyParagraphLines` →
 * `renderParagraph` in renderer.ts) still constructs segments (bidi/tab detection) and,
 * at a paint scale ≠ 1, re-measures each stored line's glyph geometry via
 * `rescaleLayoutLines` — by design (the fragment stores the scale-1 line PARTITION; only
 * glyph metrics are re-derived at the display scale, never the wrap points). Full static
 * enforcement across the whole body-paint path lands when body paint is separated into
 * its own module (PR 6+ scope).
 *
 * See docs/docx-layout-context-fragments-design.md §"Measured Fragment Model".
 */
import {
  renderBodyParagraphLines,
  renderTableFragment,
  type ParaBorderMerge,
  type RenderState,
} from './renderer.js';
import type { PlacedFragment } from './layout-fragments.js';

/**
 * Paint a body paragraph fragment at the current render cursor, drawing its stored
 * lines (rescaled) without remeasuring.
 *
 * @param placed  the placed paragraph fragment (its `measured` holds the scale-1 lines).
 * @param state   the page paint state (its `scale`, `ctx`, `contentX/W`, `y` cursor).
 * @param options paint-adjacency inputs the paginator does not own: `suppressSpaceBefore`
 *   (continuation slices / spacing collapse), the §17.3.1.7 `borderMerge`, and
 *   `continuesParagraph` (§17.6.4 remainder re-wrap: this fragment's partition is a
 *   paragraph CONTINUATION even when it covers the whole re-measured partition, so
 *   first-slice-only paint work — anchor drawing, the top border edge — must not run).
 */
export function paintParagraphFragment(
  placed: PlacedFragment,
  state: RenderState,
  options: {
    suppressSpaceBefore?: boolean;
    borderMerge?: ParaBorderMerge;
    continuesParagraph?: boolean;
  } = {},
): void {
  const fragment = placed.fragment;
  // Body paragraph paint only. A table fragment is painted by the table-fragment
  // path; the migration gate ({@link isFragmentPaintableParagraph}) never routes one
  // here, so this narrow is defensive and a no-op in production.
  if (fragment.kind !== 'paragraph') return;
  const measured = fragment.measured;
  // The stored scale-1 line partition for the whole paragraph. Empty for a markOnly
  // (empty / anchor-only) paragraph — the renderer's empty-mark branch handles it.
  const scale1Lines = measured.lines.map((line) => line.layout);
  // A full-range fragment paints the whole paragraph (no slice); a continuation
  // fragment paints only its `[lineStart, lineEnd)` window. A RE-WRAPPED remainder
  // (issue #908) covers its whole re-measured partition yet is still a continuation:
  // keep the slice object with `continues` so the draw path mirrors the legacy
  // element-slice semantics exactly (no re-run of first-slice-only work).
  const continues = options.continuesParagraph === true;
  const lineSlice =
    fragment.lineStart === 0 && fragment.lineEnd === measured.lines.length && !continues
      ? undefined
      : {
          start: fragment.lineStart,
          end: fragment.lineEnd,
          ...(continues ? { continues: true as const } : {}),
        };
  renderBodyParagraphLines(
    fragment.source,
    state,
    scale1Lines,
    options.suppressSpaceBefore ?? false,
    lineSlice,
    options.borderMerge,
  );
}

/**
 * Paint a block-table {@link PlacedFragment} produced by {@link layoutDocument} / the
 * paginator: the table's geometry (column widths, row heights) and cell content
 * (paragraph + nested-table fragments) are drawn from the stored {@link TableFragment}
 * WITHOUT re-laying-out or re-measuring. At the paint scale of 1 no `measureText` is
 * called at all. Like {@link paintParagraphFragment}, this owns the fragment boundary
 * and delegates the draw to the renderer's shared table paint through a supplied-geometry
 * entry (`renderTableFragment`), so the paint stream is byte-identical to the legacy
 * `renderTable` acquisition for the migrated in-flow table class, including supported
 * center/bottom-aligned and sliced cell fragments.
 */
export function paintTableFragment(placed: PlacedFragment, state: RenderState): void {
  const fragment = placed.fragment;
  // Table fragment paint only; the migration gate never routes a paragraph fragment
  // here, so this narrow is defensive and a no-op in production.
  if (fragment.kind !== 'table') return;
  renderTableFragment(fragment, state);
}
