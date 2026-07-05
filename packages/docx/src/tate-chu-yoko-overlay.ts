import type { DocxTextRunInfo } from './renderer';

/**
 * ECMA-376 §17.3.2.10 eastAsianLayout `w:vert` (縦中横 / horizontal-in-vertical)
 * overlay geometry — shared by the find-highlight overlay and the text-selection
 * overlay so both clamp a tate-chu-yoko run identically (packages/docx/CLAUDE.md
 * — no duplicated geometry).
 *
 * The 縦中横 draw path lays a run's glyphs out horizontally, side by side,
 * COMPRESSED into ONE em cell of the vertical column (see `drawTateChuYokoRun`):
 * the digits "２９" whose natural advance is ~2 em are drawn inside a single
 * ~1 em cell. `onTextRun` reports that drawn cell as `run.w` (pinned to one em by
 * `segAdvanceWidth`'s 縦中横 branch), but the overlays otherwise size themselves
 * from the run's NATURAL glyph width — so a find / selection box over "２９"
 * overshoots ~2× into the following cell (#836).
 *
 * The correct clamp is a single horizontal SCALE: the drawn run is exactly its
 * natural glyphs compressed by `run.w / naturalWidth` (this folds in `w:w`, which
 * further narrows the digits — the compression is whatever it takes to land the
 * natural glyphs inside the reported cell). Applying that one factor keeps every
 * intra-run offset proportional, so a partial match / partial selection inside the
 * cell still maps to the right sub-extent.
 *
 * @param run      the run to test.
 * @param measure  advance width (px) of a substring in the run's font — a
 *                 `ctx.measureText(s).width` closure already primed with
 *                 `run.font`. Only called for a 縦中横 run.
 * @returns the horizontal scale factor to apply to the run's natural overlay
 *          extent, or `1` when the run is not 縦中横 (or its natural width is
 *          degenerate / already ≤ the cell, so no compression is needed).
 */
export function tateChuYokoOverlayScale(
  run: DocxTextRunInfo,
  measure: (s: string) => number,
): number {
  if (!run.eastAsianVert) return 1;
  const natural = measure(run.text);
  // Guard a zero / NaN measure (headless without canvas metrics) and never
  // EXPAND — the cell is the ceiling; a run already within it keeps scale 1.
  if (!(natural > 0) || run.w >= natural) return 1;
  return run.w / natural;
}
