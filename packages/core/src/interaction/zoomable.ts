/**
 * IX9 — the shared zoom API contract for every viewer (DocxViewer, PptxViewer,
 * DocxScrollViewer, PptxScrollViewer, XlsxViewer).
 *
 * This module owns ONLY the pure, DOM-free pieces of the contract: the type
 * ({@link ZoomableViewer}), the discrete zoom-step ladder ({@link nextZoomStep} /
 * {@link prevZoomStep}), the fit-to-content scale math ({@link fitScale}), and the
 * range clamp ({@link clampScale}). Each viewer implements the interface with its
 * own scale field and re-render path; this keeps ONE definition of "what a zoom
 * factor means" and "what the +/- steps are" across all five, so a host can drive
 * any viewer through the same six calls without special-casing the format.
 *
 * SCALE SEMANTICS (the contract): a scale of `1` means 100% — the natural /
 * fit-to-width baseline. `getScale()` and `setScale(n)` speak this user-facing
 * factor for EVERY viewer, even though the viewers store scale differently
 * internally (XlsxViewer's `cellScale` already is this factor; the single-canvas
 * viewers multiply the page's natural pt→px size; the scroll viewers multiply
 * their established fit-to-width base). Implementations translate to/from their
 * internal representation at the boundary so the contract stays uniform.
 *
 * API SHAPE (idiomatic default — the integrator MAY veto; see the IX9 PR): a
 * six-method surface plus one change notification (`onScaleChange`). Deliberately
 * NO new UI here — the contract is API only (design decision IX9 §4). Touch-pinch
 * (IX8) is out of scope.
 */

/**
 * The zoom contract every viewer satisfies. All scales are the user-facing factor
 * where `1` = 100% (see the module note). `fitWidth`/`fitPage` are async because a
 * fit re-renders at the new scale; the getters/steppers resolve synchronously.
 */
export interface ZoomableViewer {
  /** The current zoom factor (`1` = 100%). Never throws — returns the default
   *  (`1`) before anything is loaded. */
  getScale(): number;
  /** Set the absolute zoom factor (`1` = 100%), clamped to the viewer's
   *  `[zoomMin, zoomMax]`. Re-renders at the new scale and fires `onScaleChange`
   *  when the clamped value actually changes. */
  setScale(scale: number): void | Promise<void>;
  /** Step up to the next larger rung of the shared zoom ladder (25 %→400 %),
   *  clamped to `zoomMax`. Equivalent to `setScale(nextZoomStep(getScale()))`. */
  zoomIn(): void | Promise<void>;
  /** Step down to the next smaller ladder rung, clamped to `zoomMin`. */
  zoomOut(): void | Promise<void>;
  /** Fit the content's WIDTH to the container (the common "fit width" / "fit
   *  page width" verb). Sets the scale so one page/slide/sheet-column-run spans
   *  the available width, then re-renders. Resolves once the fit render settles. */
  fitWidth(): void | Promise<void>;
  /** Fit the WHOLE content (width AND height) inside the container, so an entire
   *  page/slide is visible without scrolling. Sets the scale to the smaller of the
   *  width- and height-fit factors, then re-renders. */
  fitPage(): void | Promise<void>;
}

/**
 * The discrete zoom rungs the +/- steppers snap through, as user-facing factors.
 * The Excel / Acrobat family (25, 33, 50, 67, 75, 90, 100, 110, 125, 150, 175,
 * 200, 250, 300, 400 %) — one consistent ascending series (a superset of Excel's
 * status-bar presets and Acrobat's toolbar presets), so `zoomIn`/`zoomOut` feel
 * familiar in either lineage. Frozen (readonly) so a caller cannot mutate the
 * shared ladder. Values are factors, not percents (0.25 … 4).
 */
export const ZOOM_STEP_LADDER: readonly number[] = Object.freeze([
  0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4,
]);

/** Small epsilon so a `scale` sitting exactly on (or a floating-point hair away
 *  from) a ladder rung steps to the NEXT rung, not back to the one it is already
 *  on. 0.5 % of a factor is far below the gap between any two adjacent rungs. */
const LADDER_EPS = 0.005;

/**
 * The next ladder rung strictly greater than `scale` (the "+" step). A `scale`
 * already at or above the top rung returns the top rung (the caller then clamps
 * to `zoomMax`, which may be higher or lower than the ladder top). A `scale`
 * between rungs jumps UP to the next rung — so an arbitrary wheel-zoomed value
 * snaps back onto the ladder on the first "+" press. Pure.
 */
export function nextZoomStep(scale: number): number {
  for (const rung of ZOOM_STEP_LADDER) {
    if (rung > scale + LADDER_EPS) return rung;
  }
  return ZOOM_STEP_LADDER[ZOOM_STEP_LADDER.length - 1];
}

/**
 * The next ladder rung strictly less than `scale` (the "−" step). Mirror of
 * {@link nextZoomStep}: a `scale` at or below the bottom rung returns the bottom
 * rung; a between-rungs value jumps DOWN to the next rung. Pure.
 */
export function prevZoomStep(scale: number): number {
  for (let i = ZOOM_STEP_LADDER.length - 1; i >= 0; i--) {
    const rung = ZOOM_STEP_LADDER[i];
    if (rung < scale - LADDER_EPS) return rung;
  }
  return ZOOM_STEP_LADDER[0];
}

/** Clamp `scale` to `[min, max]`. Pure; `max < min` yields `min` (degenerate
 *  bounds resolve to the floor, matching the viewers' inline clamps). */
export function clampScale(scale: number, min: number, max: number): number {
  return scale < min ? min : scale > max ? max : scale;
}

/** Content + container extents (px, same axis) for a fit computation. */
export interface FitInput {
  /** Natural content width at 100% (CSS px). */
  contentWidth: number;
  /** Natural content height at 100% (CSS px). */
  contentHeight: number;
  /** Available container width (CSS px). */
  containerWidth: number;
  /** Available container height (CSS px). Only used by {@link fitScale} in
   *  `'page'` mode; a `'width'` fit ignores it. */
  containerHeight: number;
}

/** Which dimension a fit targets: `'width'` fits the width only (height scrolls);
 *  `'page'` fits both so the whole page is visible. */
export type FitMode = 'width' | 'page';

/**
 * The user-facing zoom factor that fits the content to the container.
 *
 * - `'width'`: `containerWidth / contentWidth` — the page spans the width, height
 *   overflows into the scroll region.
 * - `'page'`: `min(containerWidth / contentWidth, containerHeight / contentHeight)`
 *   — the whole page fits inside the box (letterboxed on the looser axis).
 *
 * Returns `0` when any input needed for the chosen mode is non-positive (an
 * unlaid-out container or empty content) so the caller can DEFER — the same
 * zero-as-defer convention the scroll viewers already use for their base fit.
 * Pure; no DOM, no clamping (the caller clamps to `[zoomMin, zoomMax]`).
 */
export function fitScale(input: FitInput, mode: FitMode): number {
  const { contentWidth, contentHeight, containerWidth, containerHeight } = input;
  if (contentWidth <= 0 || containerWidth <= 0) return 0;
  const widthFit = containerWidth / contentWidth;
  if (mode === 'width') return widthFit;
  if (contentHeight <= 0 || containerHeight <= 0) return 0;
  const heightFit = containerHeight / contentHeight;
  return Math.min(widthFit, heightFit);
}
