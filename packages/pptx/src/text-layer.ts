import type { PptxTextRunInfo } from './renderer';

/**
 * Build the transparent text-selection overlay for a rendered pptx slide. Unlike
 * docx (flat spans), pptx groups runs into one positioned `<div>` per shape frame
 * (keyed by the shape's geometry + total rotation) and applies a CSS `rotate()` to
 * the group when the shape is rotated, so the browser selection tracks the drawn,
 * rotated text as a unit. Each run's `<span>` is absolutely positioned INSIDE its
 * shape div (`inShapeX`/`inShapeY`). Extracted verbatim from
 * `PptxViewer._buildTextLayer` so the pager (PptxViewer) and the continuous-scroll
 * viewer (PptxScrollViewer, WS4) share one implementation; public API for
 * integrators (design §10). MAIN render mode only — `onTextRun` cannot cross the
 * worker boundary.
 *
 * @param layer     the overlay div.
 * @param runs      per-run + per-shape geometry from `renderSlide({ onTextRun })`.
 * @param cssWidth  the rendered canvas's CSS width (px, number).
 * @param cssHeight the rendered canvas's CSS height (px, number).
 */
export function buildPptxTextLayer(
  layer: HTMLDivElement,
  runs: PptxTextRunInfo[],
  cssWidth: number,
  cssHeight: number,
): void {
  layer.innerHTML = '';
  layer.style.width = `${cssWidth}px`;
  layer.style.height = `${cssHeight}px`;

  // Group runs by shape (same shapeX/shapeY/rotation)
  type ShapeKey = string;
  const shapeMap = new Map<ShapeKey, { div: HTMLDivElement; x: number; y: number; w: number; h: number; rot: number }>();

  for (const run of runs) {
    const totalRot = run.rotation + (run.textBodyRotation ?? 0);
    const key = `${run.shapeX},${run.shapeY},${run.shapeW},${run.shapeH},${totalRot}`;
    if (!shapeMap.has(key)) {
      const div = document.createElement('div');
      div.style.cssText =
        `position:absolute;` +
        `left:${run.shapeX}px;top:${run.shapeY}px;` +
        `width:${run.shapeW}px;height:${run.shapeH}px;` +
        `pointer-events:all;overflow:hidden;`;
      if (totalRot !== 0) {
        div.style.transformOrigin = 'center center';
        div.style.transform = `rotate(${totalRot}deg)`;
      }
      shapeMap.set(key, { div, x: run.shapeX, y: run.shapeY, w: run.shapeW, h: run.shapeH, rot: totalRot });
      layer.appendChild(div);
    }

    const shape = shapeMap.get(key)!;
    const span = document.createElement('span');
    span.textContent = run.text;
    // The `font` shorthand must precede `line-height` because the shorthand
    // resets `line-height` to `normal`. Reset `letter-spacing` so a parent
    // CSS rule cannot drift the trailing edge of the selection. Kerning /
    // ligatures are left at the browser default ('auto') because canvas
    // `measureText` / `fillText` also apply them by default — forcing them
    // off here would make the span wider than the drawn text.
    span.style.cssText =
      `position:absolute;` +
      `left:${run.inShapeX}px;top:${run.inShapeY}px;` +
      `font:${run.font};line-height:${run.h}px;letter-spacing:0;` +
      `white-space:pre;color:transparent;cursor:text;`;
    shape.div.appendChild(span);
  }
}
