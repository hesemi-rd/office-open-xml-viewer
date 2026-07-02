/**
 * Shared vector-vs-raster blip selection gate for the docx / pptx / xlsx
 * renderers.
 *
 * A DrawingML picture can carry a Microsoft `asvg:svgBlip` extension
 * (MS-ODRAWXML) alongside the ordinary raster `<a:blip>`: a vector original that
 * we prefer over the raster fallback for crisp scaling. The one exception is a
 * cropped picture — an `<a:srcRect>` source rectangle (ECMA-376 §20.1.8.55)
 * expresses the crop as FRACTIONS of the source's pixel grid, so the crop math
 * (`drawImageCropped`) needs the decoded bitmap's native pixel dimensions. An
 * SVG `HTMLImageElement` that declares only a `viewBox` (no intrinsic
 * width/height) reports the 300×150 default rather than its logical size, which
 * would make a 9-argument `drawImage` with a source rect sample the wrong basis.
 * So when a crop is present we skip the vector original and decode the raster,
 * whose exact pixel dimensions make the fractional crop well-defined.
 *
 * The argument is a structural type so each format's own picture/image model
 * assigns cleanly. `srcRect` is treated purely by PRESENCE: any non-nullish
 * value means "cropped" (callers that only retain an aggregated boolean crop
 * flag pass `flag || null`).
 *
 * A type guard: when it returns `true`, the caller's `svgImagePath` is narrowed
 * to a non-null `string`, so the vector decode can pass it straight to
 * `getCachedSvgImageByPath` without a re-check.
 *
 * @returns `true` to prefer the vector original (has an svgImagePath, uncropped);
 *   `false` to use the raster path (no svgImagePath, or a crop is present).
 */
export function preferVectorBlip<T extends { svgImagePath?: string | null; srcRect?: unknown }>(
  el: T,
): el is T & { svgImagePath: string } {
  return el.svgImagePath != null && el.srcRect == null;
}
