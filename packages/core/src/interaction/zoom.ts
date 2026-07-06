/** Ctrl/⌘ + wheel (and trackpad pinch) zoom sensitivity. `deltaY` is multiplied
 *  by this before `exp()`. Purely an interaction-feel constant (no ECMA-376
 *  bearing); lower = gentler. */
const ZOOM_WHEEL_SENSITIVITY = 0.01;

/**
 * New scale for one wheel/pinch zoom step. The step is *exponential* in
 * `deltaY` rather than a fixed increment, which fixes two problems with a
 * sign-only `scale ± 0.1`:
 *
 *  - A trackpad pinch arrives as a high-frequency stream of small-`deltaY`
 *    wheel events; a fixed per-event increment compounds across dozens of
 *    events per gesture and zooms wildly. Because `exp(-k·a)·exp(-k·b) =
 *    exp(-k·(a+b))`, the total zoom here depends only on the summed `deltaY`
 *    of the gesture, not on how many events the OS chops it into — so a pinch
 *    and a mouse wheel covering the same distance zoom by the same amount.
 *  - It is multiplicative, so a step feels proportional at every zoom level
 *    (the old additive `+0.1` was huge at 20% and tiny at 400%), and exactly
 *    symmetric: zooming in then out by the same delta returns to the start.
 *
 * Negative `deltaY` (scroll up / pinch out) zooms in. The result is unclamped
 * and unsnapped; the caller clamps to its `[zoomMin, zoomMax]` range (and, for
 * XlsxViewer, snaps to whole percent). Shared by XlsxViewer and the two
 * continuous-scroll viewers (design §5.2).
 */
export function zoomStepScale(currentScale: number, deltaY: number): number {
  return currentScale * Math.exp(-deltaY * ZOOM_WHEEL_SENSITIVITY);
}

/** Clamp range for {@link anchoredZoomOffset}: the scroll offset must stay within
 *  `[0, maxScroll]` (the browser clamps native scrollLeft/scrollTop the same way).
 *  `maxScroll` is the largest legal scroll on that axis at the NEW scale (usually
 *  `totalContentPx − viewportPx`, floored at 0). */
export interface ScrollClamp {
  /** Largest legal scroll offset on this axis at the new scale (≥ 0). A degenerate
   *  `maxScroll < 0` (content shorter than the viewport) is treated as 0. */
  maxScroll: number;
}

/**
 * The new scroll offset (one axis) that keeps the content point currently under a
 * gesture anchor pinned under that same anchor across a zoom — i.e. pointer-
 * anchored ("zoom toward the cursor") zoom, instead of the viewport-top / origin
 * anchoring a naive rescale gives.
 *
 * DERIVATION. Work in the SCALING content's own pixels, both measured from the
 * start of that region (the caller subtracts any scale-invariant lead-in — desk
 * padding, a frozen header/pane — from the pointer BEFORE calling, so `anchor` is
 * the pointer's distance into the region that actually zooms):
 *
 *   - `scrollOld` — current scroll offset into the region (scaled px, ≥ 0).
 *   - `anchor`    — the pointer's offset from the region's on-screen start
 *                   (scaled px). The content pixel under the pointer is therefore
 *                   `c = scrollOld + anchor`.
 *   - A zoom multiplies every content pixel by `ratio = scaleNew / scaleOld`, so
 *     that SAME logical content sits at `c' = c · ratio` at the new scale.
 *   - To keep it under the (unchanged, screen-fixed) pointer, we need
 *     `scrollNew + anchor = c'`, hence:
 *
 *         scrollNew = (scrollOld + anchor) · (scaleNew / scaleOld) − anchor
 *
 * IDENTITY: `scaleNew === scaleOld` ⇒ `ratio = 1` ⇒ `scrollNew = scrollOld` (no
 * drift on a no-op zoom). ROUND-TRIP: zooming out by the inverse ratio returns the
 * original offset exactly (the map is affine and invertible). ANCHOR = 0 (pointer
 * at the region start) reduces to the plain `scrollOld · ratio` origin-anchored
 * rescale.
 *
 * CLAMP: the result is clamped to `[0, clamp.maxScroll]` — the same range the
 * browser pins native scroll to. AT A CONTENT EDGE the anchor cannot always be
 * honoured (there is no scroll offset that would place the pinned point under the
 * pointer without exposing off-content area); the clamp then wins and the point
 * drifts, which is the natural, expected behaviour (design §5.3).
 *
 * Pure; no DOM. Shared by the two continuous-scroll viewers (both axes) and
 * XlsxViewer (both axes, past its fixed header + frozen panes).
 */
export function anchoredZoomOffset(
  scrollOld: number,
  anchor: number,
  scaleOld: number,
  scaleNew: number,
  clamp: ScrollClamp,
): number {
  // Guard a degenerate old scale (never happens for a live viewer, but keeps the
  // function total): fall back to no movement rather than dividing by zero.
  const ratio = scaleOld > 0 ? scaleNew / scaleOld : 1;
  const want = (scrollOld + anchor) * ratio - anchor;
  const max = clamp.maxScroll > 0 ? clamp.maxScroll : 0;
  return want < 0 ? 0 : want > max ? max : want;
}
