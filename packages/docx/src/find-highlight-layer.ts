/**
 * IX2 docx find-highlight overlay.
 *
 * The mirror image of {@link buildDocxTextLayer}: instead of a transparent
 * selection span per run, it draws a visible highlight box per matched
 * run-slice, positioned with the SAME per-run geometry (`DocxTextRunInfo.x/y/h`
 * + the run's `transform` for vertical pages) so a box lands exactly on the
 * drawn glyphs. This rides the existing DOM-overlay mechanism (a positioned
 * layer over the canvas) rather than adding a canvas draw pass — the same
 * approach the selection / hyperlink overlay uses, so highlights compose with
 * them and survive re-render the same way.
 *
 * The horizontal extent of a slice (its x offset + width within the run) is the
 * shared core `sliceHorizontalExtent`, measured against the run's font via a
 * caller-supplied `measure` primed with `ctx.font = run.font` — the same
 * measureText↔fillText correspondence the selection overlay relies on. The
 * vertical extent is the run's line box (`y` … `y + h`), owned here.
 *
 * The active match (the one `findNext`/`findPrev` last landed on) is drawn in a
 * distinct emphasis colour so the user can tell it apart from the other hits,
 * matching a browser find bar's current-vs-other highlighting.
 */
import { sliceHorizontalExtent, type MatchRunSlice } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer';
import { tateChuYokoOverlayScale } from './tate-chu-yoko-overlay';

/** One page's highlight input: the run-slices a match covers, and whether that
 *  match is the active one (emphasis colour). */
export interface DocxHighlightMatch {
  slices: MatchRunSlice[];
  active: boolean;
}

/** Default highlight colours (browser find-bar palette): a soft yellow for
 *  every match and a stronger orange for the active one. Both translucent so the
 *  drawn glyphs beneath stay legible. */
export const DEFAULT_FIND_HIGHLIGHT = 'rgba(255, 214, 0, 0.42)';
export const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'rgba(255, 140, 0, 0.55)';

export interface DocxHighlightColors {
  /** Fill for non-active matches. */
  match?: string;
  /** Fill for the active match. */
  active?: string;
}

/**
 * Populate a highlight overlay layer with one box per matched run-slice.
 *
 * @param layer    the overlay div (cleared and re-sized to the canvas here).
 * @param runs     the page's runs (same array the page was rendered/text-layered from).
 * @param matches  the page's matches (run-slices + active flag).
 * @param canvasCssWidth  the rendered canvas's CSS width (e.g. `"700px"`).
 * @param canvasCssHeight the rendered canvas's CSS height.
 * @param measureForFont  returns a width-measurer primed with a run's `font`
 *                        (the viewer closes over a canvas 2d context). Kept as a
 *                        factory so the font is set once per run, not per glyph.
 * @param colors   optional colour overrides.
 */
export function buildDocxHighlightLayer(
  layer: HTMLDivElement,
  runs: DocxTextRunInfo[],
  matches: DocxHighlightMatch[],
  canvasCssWidth: string,
  canvasCssHeight: string,
  measureForFont: (font: string) => (s: string) => number,
  colors: DocxHighlightColors = {},
): void {
  layer.innerHTML = '';
  layer.style.width = canvasCssWidth;
  layer.style.height = canvasCssHeight;

  const matchColor = colors.match ?? DEFAULT_FIND_HIGHLIGHT;
  const activeColor = colors.active ?? DEFAULT_FIND_ACTIVE_HIGHLIGHT;

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      const measure = measureForFont(run.font);
      const extent = sliceHorizontalExtent(run.text, slice.start, slice.end, measure);
      // ECMA-376 §17.3.2.10 縦中横 (#836): a tate-chu-yoko run is drawn compressed
      // into one em cell (`run.w`), so its natural per-glyph extents overshoot the
      // drawn box. Scale the slice offset + width by `run.w / naturalWidth` so the
      // highlight lands on the compressed cell (a no-op factor of 1 for every
      // ordinary run — see tate-chu-yoko-overlay.ts). Applying one factor keeps a
      // partial match proportional within the cell.
      const k = tateChuYokoOverlayScale(run, measure);
      const x = extent.x * k;
      const width = extent.width * k;
      if (width <= 0) continue;
      const box = document.createElement('div');
      // A vertical (tbRl) page reports the run at its physical top-left and
      // carries a rotate transform; apply the same transform (about the box's
      // top-left) so the highlight lies along the rotated glyph run, exactly as
      // the selection span does.
      const transform = run.transform
        ? `transform:${run.transform};transform-origin:top left;`
        : '';
      box.style.cssText =
        `position:absolute;` +
        `left:${run.x + x}px;top:${run.y}px;` +
        `width:${width}px;height:${run.h}px;` +
        transform +
        `background:${fill};pointer-events:none;`;
      layer.appendChild(box);
    }
  }
}
