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
 * See docs/docx-layout-context-fragments-design.md §"Measured Fragment Model".
 */
import { renderBodyParagraphLines, type ParaBorderMerge, type RenderState } from './renderer.js';
import type { PlacedFragment } from './layout-fragments.js';

/**
 * Paint a body paragraph fragment at the current render cursor, drawing its stored
 * lines (rescaled) without remeasuring.
 *
 * @param placed  the placed paragraph fragment (its `measured` holds the scale-1 lines).
 * @param state   the page paint state (its `scale`, `ctx`, `contentX/W`, `y` cursor).
 * @param options paint-adjacency inputs the paginator does not own: `suppressSpaceBefore`
 *   (continuation slices / spacing collapse) and the §17.3.1.7 `borderMerge`.
 */
export function paintParagraphFragment(
  placed: PlacedFragment,
  state: RenderState,
  options: {
    suppressSpaceBefore?: boolean;
    borderMerge?: ParaBorderMerge;
  } = {},
): void {
  const fragment = placed.fragment;
  const measured = fragment.measured;
  // The stored scale-1 line partition for the whole paragraph. Empty for a markOnly
  // (empty / anchor-only) paragraph — the renderer's empty-mark branch handles it.
  const scale1Lines = measured.lines.map((line) => line.layout);
  // A full-range fragment paints the whole paragraph (no slice); a continuation
  // fragment paints only its `[lineStart, lineEnd)` window.
  const lineSlice =
    fragment.lineStart === 0 && fragment.lineEnd === measured.lines.length
      ? undefined
      : { start: fragment.lineStart, end: fragment.lineEnd };
  renderBodyParagraphLines(
    fragment.source,
    state,
    scale1Lines,
    options.suppressSpaceBefore ?? false,
    lineSlice,
    options.borderMerge,
  );
}
