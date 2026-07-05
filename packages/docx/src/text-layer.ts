import type { DocxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';
import { tateChuYokoOverlayScale } from './tate-chu-yoko-overlay';

/**
 * Build the transparent text-selection overlay for a rendered docx page: one
 * absolutely-positioned, color-transparent `<span>` per {@link DocxTextRunInfo}
 * (emitted by `renderPage`'s `onTextRun`), so the browser's native selection
 * lands on the drawn glyphs. Extracted verbatim from `DocxViewer._buildTextLayer`
 * so both the pager (DocxViewer) and the continuous-scroll viewer (DocxScrollViewer)
 * share one implementation; also public API for integrators building their own
 * overlay (design §10). IX6 — usable in BOTH render modes: worker mode collects
 * the same `DocxTextRunInfo[]` off-thread and ships it back beside the bitmap, so
 * the overlay is built from identical geometry regardless of thread.
 *
 * @param layer            the overlay div (position:relative parent expected).
 * @param runs             per-run geometry from `renderPage({ onTextRun })`.
 * @param canvasCssWidth   the rendered canvas's CSS width (e.g. `"700px"`), used
 *                         to size the overlay to match the canvas.
 * @param canvasCssHeight  the rendered canvas's CSS height.
 * @param onHyperlinkClick IX1 — invoked when a run carrying a resolved
 *                         {@link HyperlinkTarget} is clicked. A hyperlink run's
 *                         span keeps its transparent glyphs (the visible link
 *                         colour/underline is already drawn on the canvas) but
 *                         gains `cursor:pointer`, a `title` tooltip (the URL or
 *                         bookmark ref) and this click handler. A plain
 *                         `<span>` — not an `<a href>` — is used deliberately so
 *                         the browser's own navigation can never bypass the
 *                         caller's URL sanitisation. When omitted, link runs are
 *                         rendered exactly like plain runs (no click affordance).
 * @param measureForFont   optional width-measurer factory (primed with a run's
 *                         `font`), used ONLY to clamp a §17.3.2.10 縦中横
 *                         (eastAsianVert) span to its drawn one-em cell (#836):
 *                         the span composes a `scaleX(run.w / naturalWidth)` so
 *                         its selection extent matches the compressed glyphs
 *                         instead of the run's natural ~2× width. When omitted,
 *                         a 縦中横 span keeps the bare rotate (no regression for
 *                         callers that do not thread a measurer).
 */
export function buildDocxTextLayer(
  layer: HTMLDivElement,
  runs: DocxTextRunInfo[],
  canvasCssWidth: string,
  canvasCssHeight: string,
  onHyperlinkClick?: (target: HyperlinkTarget) => void,
  measureForFont?: (font: string) => (s: string) => number,
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
    // ECMA-376 §17.3.2.10 縦中横 (#836): a tate-chu-yoko run is drawn compressed
    // into one em cell (`run.w`), so append a `scaleX(run.w / naturalWidth)` (when
    // a measurer is available) so the span's selectable extent matches the drawn
    // cell instead of the natural ~2× glyph width. Ordered AFTER the rotate so it
    // compresses the span's horizontal (pre-rotation) axis — the along-column
    // direction of the cell. A no-op factor of 1 for every ordinary run (see
    // tate-chu-yoko-overlay.ts), so a non-縦中横 span's transform is unchanged.
    let transformValue = run.transform ?? '';
    if (measureForFont && run.eastAsianVert) {
      const k = tateChuYokoOverlayScale(run, measureForFont(run.font));
      if (k !== 1) transformValue = `${transformValue ? `${transformValue} ` : ''}scaleX(${k})`;
    }
    const transform = transformValue
      ? `transform:${transformValue};transform-origin:top left;`
      : '';
    // IX1 — a run with a resolved hyperlink target becomes a clickable region.
    // Only the cursor changes to a pointer and a title tooltip is added; the
    // glyphs stay `color:transparent` (the link's blue/underline is already
    // painted on the canvas), so a link run's selection behaviour is otherwise
    // identical to a plain run. A non-link run is byte-identical to before.
    const link = onHyperlinkClick ? run.hyperlink : undefined;
    const cursor = link ? 'pointer' : 'text';
    span.style.cssText =
      `position:absolute;` +
      `left:${run.x}px;top:${run.y}px;` +
      `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
      transform +
      `white-space:pre;color:transparent;cursor:${cursor};pointer-events:all;`;
    if (link && onHyperlinkClick) {
      span.title = link.kind === 'external' ? link.url : link.ref;
      span.addEventListener('click', () => onHyperlinkClick(link));
    }
    layer.appendChild(span);
  }
}
