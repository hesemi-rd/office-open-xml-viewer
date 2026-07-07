/**
 * Turn a matched run-slice (from {@link findMatches}) into the horizontal extent
 * of the highlight box within its run, using a caller-supplied text-measurer.
 *
 * A run's glyphs are laid left-to-right from the run's own origin. A slice
 * `[start, end)` of the run text therefore starts at the advance width of
 * `runText[0..start)` and ends at the advance width of `runText[0..end)`. The
 * caller supplies `measure(s)` — a `CanvasRenderingContext2D.measureText(s).width`
 * closure already primed with the run's font (`ctx.font = run.font`) — so this
 * stays a pure arithmetic wrapper with no canvas/DOM dependency of its own and is
 * shared by the docx and pptx highlight overlays (xlsx measures inside the cell
 * rect the same way). The vertical extent (top / height) is the run's line box,
 * owned by each renderer's overlay, so it is not computed here.
 *
 * Measuring the two prefixes (rather than measuring the slice text alone and
 * summing) is deliberate: it accounts for kerning between the run's leading
 * glyphs and keeps the highlight edges flush with where the canvas actually drew
 * those characters, exactly as the selection overlay relies on `measureText`
 * matching `fillText`.
 *
 * @param runText   the full text of the run the slice belongs to.
 * @param start     slice start offset within `runText` (inclusive).
 * @param end       slice end offset within `runText` (exclusive).
 * @param measure   advance width in px of a substring, in the run's font.
 * @returns `x` (left offset from the run origin, px) and `width` (px).
 */
export function sliceHorizontalExtent(
  runText: string,
  start: number,
  end: number,
  measure: (s: string) => number,
): { x: number; width: number } {
  const x = start <= 0 ? 0 : measure(runText.slice(0, start));
  const endX = end >= runText.length ? measure(runText) : measure(runText.slice(0, end));
  return { x, width: Math.max(0, endX - x) };
}

/**
 * Express an overlay coordinate as a CSS **percentage** of the content box it is
 * measured against, e.g. `overlayPercent(480, 960) === '50%'`.
 *
 * The find-highlight / text-selection / hyperlink overlays (docx + pptx) place
 * their boxes over a `<canvas>` that a consumer may scale down with external CSS
 * (`width:100%!important; height:auto`). If the overlay's coordinates were literal
 * px (the slide's *intended* size) while the canvas rendered smaller, the boxes
 * would overshoot the wrapper and push a scrollbar onto an ancestor's scroll area
 * (the reported bug). Positioning every box as a `%` of the intended CSS box
 * (`cssWidth`/`cssHeight`) instead makes it resolve against the container's
 * ACTUAL rendered size — the wrapper is `display:inline-block`, so it tracks the
 * scaled canvas, and each `%`-placed child scales with it. At 1× (no external
 * scaling) the box lands on exactly the same pixels as the old px placement.
 *
 * A `0`/negative denominator (nothing laid out yet) yields `'0%'` so the caller
 * never emits `NaN%`.
 *
 * @param value  the coordinate/length in the intended CSS-px space.
 * @param basis  the intended CSS-px extent it is a fraction of (`cssWidth` for the
 *               x axis, `cssHeight` for the y axis).
 */
export function overlayPercent(value: number, basis: number): string {
  if (!(basis > 0)) return '0%';
  return `${(value / basis) * 100}%`;
}
