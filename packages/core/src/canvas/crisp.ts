/**
 * Crisp-line rasterization helper shared by the docx / xlsx / pptx renderers.
 *
 * All three renderers draw in *logical* pixels under `ctx.scale(dpr, dpr)`, so a
 * thin axis-aligned stroke can land between device rows and render blurry (its
 * ink spread across two rows at partial coverage). `crispOffset` returns the
 * delta to add to an axis-aligned stroke's coordinate so it snaps onto the
 * nearest crisp device position.
 *
 * Given the stroke's logical `coord` (the x of a vertical line / the y of a
 * horizontal line) and `logicalWidth`, the device-space width is
 * `round(logicalWidth * dpr)`. An ODD device width must sit on a pixel *midpoint*
 * to occupy a single device row; an EVEN device width must sit on an integer
 * device boundary. We snap `coord` to the nearest such position and return the
 * (sub-pixel) delta — so the result is correct whether `coord` is integer- or
 * fractional-aligned, and is `0` for an even width already on the boundary.
 *
 * Only meaningful for axis-aligned (horizontal / vertical) strokes — diagonals,
 * arbitrary paths and glyph outlines cannot be pixel-aligned this way.
 */
export function crispOffset(coord: number, logicalWidth: number, dpr: number): number {
  // Target fractional part (in device px) of the stroke's center: 0.5 for odd
  // device widths (pixel midpoint), 0 for even (pixel boundary).
  const target = Math.round(logicalWidth * dpr) % 2 === 1 ? 0.5 : 0;
  const deviceCoord = coord * dpr;
  const snapped = Math.round(deviceCoord - target) + target;
  return snapped / dpr - coord;
}
