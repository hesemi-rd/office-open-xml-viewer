import type { DocxTextRunInfo } from './renderer';

/**
 * Build the transparent text-selection overlay for a rendered docx page: one
 * absolutely-positioned, color-transparent `<span>` per {@link DocxTextRunInfo}
 * (emitted by `renderPage`'s `onTextRun`), so the browser's native selection
 * lands on the drawn glyphs. Extracted verbatim from `DocxViewer._buildTextLayer`
 * so both the pager (DocxViewer) and the continuous-scroll viewer (DocxScrollViewer)
 * share one implementation; also public API for integrators building their own
 * overlay (design §10). MAIN render mode only — `onTextRun` cannot cross the
 * worker boundary.
 *
 * @param layer           the overlay div (position:relative parent expected).
 * @param runs            per-run geometry from `renderPage({ onTextRun })`.
 * @param canvasCssWidth  the rendered canvas's CSS width (e.g. `"700px"`), used
 *                        to size the overlay to match the canvas.
 * @param canvasCssHeight the rendered canvas's CSS height.
 */
export function buildDocxTextLayer(
  layer: HTMLDivElement,
  runs: DocxTextRunInfo[],
  canvasCssWidth: string,
  canvasCssHeight: string,
): void {
  layer.innerHTML = '';
  layer.style.width = canvasCssWidth;
  layer.style.height = canvasCssHeight;

  for (const run of runs) {
    const span = document.createElement('span');
    span.textContent = run.text;
    // The `font` shorthand must precede `line-height` because the shorthand
    // resets `line-height` to `normal`. Reset `letter-spacing` so a parent
    // CSS rule cannot drift the trailing edge of the selection. Kerning /
    // ligatures are left at the browser default ('auto') because canvas
    // `measureText` / `fillText` also apply them by default — forcing them
    // off here would make the span wider than the drawn text.
    // ECMA-376 §17.6.20 (tbRl) — a vertical page reports `x`/`y` as the PHYSICAL
    // top-left and carries a `transform` (`rotate(90deg)`); rotate the span about
    // its top-left so the horizontal DOM text lies along the drawn (rotated) glyph
    // run. Horizontal pages carry no transform and place the span at `x`/`y`
    // untransformed (byte-identical to the pre-vertical overlay).
    const transform = run.transform
      ? `transform:${run.transform};transform-origin:top left;`
      : '';
    span.style.cssText =
      `position:absolute;` +
      `left:${run.x}px;top:${run.y}px;` +
      `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
      transform +
      `white-space:pre;color:transparent;cursor:text;pointer-events:all;`;
    layer.appendChild(span);
  }
}
