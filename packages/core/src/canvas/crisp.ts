/**
 * Crisp-line rasterization helper shared by the docx / xlsx / pptx renderers.
 *
 * All three renderers draw in *logical* pixels under `ctx.scale(dpr, dpr)`, so a
 * thin axis-aligned stroke placed on an integer coordinate can straddle two
 * device rows (each at 50% coverage) and render blurry. Nudging the coordinate
 * by half a device pixel centers an odd-device-width stroke on a single device
 * row, which renders crisp.
 */

/**
 * The offset to add to an integer-aligned coordinate so an axis-aligned stroke of
 * `logicalWidth` logical px renders crisp under `ctx.scale(dpr, dpr)`.
 *
 * `ctx.scale(dpr, dpr)` maps the stroke width to `round(logicalWidth * dpr)`
 * device px. An ODD device width must sit on a pixel *midpoint* (offset `0.5/dpr`)
 * to land on one device row; an EVEN device width is already crisp on the integer
 * coordinate and must NOT be nudged (offset `0`), or it straddles and blurs.
 *
 * Only meaningful for axis-aligned (horizontal / vertical) strokes — diagonals,
 * arbitrary paths and glyph outlines cannot be pixel-aligned this way.
 */
export function crispOffset(logicalWidth: number, dpr: number): number {
  return Math.round(logicalWidth * dpr) % 2 === 1 ? 0.5 / dpr : 0;
}
