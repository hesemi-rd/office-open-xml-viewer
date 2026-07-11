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
import { sliceHorizontalExtent, overlayPercent, type MatchRunSlice } from '@silurus/ooxml-core';
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
 * Every box is positioned as a PERCENTAGE of `cssWidth`/`cssHeight`, and the
 * container's own size is left untouched (`width:100%;height:100%` from the
 * caller), so the highlights track the canvas's ACTUAL rendered box even when a
 * consumer scales the canvas down with external CSS — mirroring
 * {@link buildDocxTextLayer}.
 *
 * @param layer    the overlay div (cleared here; sized `100%` by the caller).
 * @param runs     the page's runs (same array the page was rendered/text-layered from).
 * @param matches  the page's matches (run-slices + active flag).
 * @param cssWidth  the page's intended CSS width (px, number) — the x-axis % denominator.
 * @param cssHeight the page's intended CSS height (px, number) — the y-axis % denominator.
 * @param measureForFont  returns a width-measurer primed with a run's `font`
 *                        (the viewer closes over a canvas 2d context). Kept as a
 *                        factory so the font is set once per run, not per glyph.
 * @param colors   optional colour overrides.
 */
export function buildDocxHighlightLayer(
  layer: HTMLDivElement,
  runs: DocxTextRunInfo[],
  matches: DocxHighlightMatch[],
  cssWidth: number,
  cssHeight: number,
  measureForFont: (font: string) => (s: string) => number,
  colors: DocxHighlightColors = {},
): void {
  layer.innerHTML = '';

  const matchColor = colors.match ?? DEFAULT_FIND_HIGHLIGHT;
  const activeColor = colors.active ?? DEFAULT_FIND_ACTIVE_HIGHLIGHT;

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      const measure = measureForFont(run.font);
      const extent = sliceHorizontalExtent(run.text, slice.start, slice.end, measure);
      // Canvas draws glyph i at measure(prefix_i) + i*pitch. Keep the shared
      // natural-width slice intact and compose the DOCX-only pitch before the
      // 縦中横 scale; no pitch follows the slice's final glyph.
      const pitch = run.letterSpacingPx ?? 0;
      const end = Math.min(slice.end, [...run.text].length);
      const pitchedX = extent.x + slice.start * pitch;
      const pitchedWidth = extent.width + Math.max(0, end - slice.start - 1) * pitch;
      // 縦中横 keeps its established one-em clamp. Ordinary horizontal runs use
      // the composed Canvas glyph scale reported by the renderer. Applying the
      // factor after pitch composition preserves the same approximation as the
      // 縦中横 path: fixed pitch is scaled with the slice rather than modelled as
      // a second coordinate axis, while offsets and widths remain proportional.
      const k = run.eastAsianVert
        ? tateChuYokoOverlayScale(run, measure)
        : (run.glyphScaleX ?? 1);
      const x = pitchedX * k;
      const width = pitchedWidth * k;
      if (width <= 0) continue;
      const box = document.createElement('div');
      // A vertical (tbRl) page reports the run at its physical top-left and
      // carries a rotate transform; apply the same transform (about the box's
      // top-left) so the highlight lies along the rotated glyph run, exactly as
      // the selection span does.
      const transform = run.transform
        ? `transform:${run.transform};transform-origin:top left;`
        : '';
      // Positioned as a % of the page's intended CSS box so the box scales with
      // the container under external CSS scaling. A vertical (tbRl) run's
      // transform is about `top left`, so the %-placed origin rotates into the
      // column exactly as before at 1× scale.
      box.style.cssText =
        `position:absolute;` +
        `left:${overlayPercent(run.x + x, cssWidth)};top:${overlayPercent(run.y, cssHeight)};` +
        `width:${overlayPercent(width, cssWidth)};height:${overlayPercent(run.h, cssHeight)};` +
        transform +
        `background:${fill};pointer-events:none;`;
      layer.appendChild(box);
    }
  }
}
