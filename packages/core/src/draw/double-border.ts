/**
 * Shared geometry for ECMA-376 `double` borders (§17.18.2 ST_Border /
 * §18.18.3 ST_BorderStyle): two parallel lines with a gap. The standard leaves
 * the rail/gap PIXEL geometry to the implementation; this renders three bands
 * across the nominal width (rail / gap / rail, each ≈ lw/3) with each band
 * FLOORED at one device pixel so a thin double (e.g. a 0.75px hairline) never
 * collapses into a single line. The floor is a rendering-legibility minimum, not
 * a content heuristic. For thick borders it reduces to equal thirds.
 *
 * Used by the docx renderer's `drawBorderLine`. Lives in core (alongside
 * `crispOffset` in canvas/crisp.ts and the dash arrays in draw/dash.ts) so any
 * fill-based double-border painter shares one source of truth.
 */

/** Device-pixel band layout of a `double` border of stroked width `lw` (px) on a
 *  `ctx.scale(dpr,dpr)`-d canvas. `railDev`/`gapDev` are ≥ 1 device pixel; the
 *  total span is `2·railDev + gapDev`. With `gapDev ≥ 1` the two rails are always
 *  separated by at least one device pixel, and with `railDev ≥ 1` each rail
 *  always paints — so the double never collapses to a single line. */
export function doubleRailGeometry(lw: number, dpr: number): {
  railDev: number;
  gapDev: number;
  spanDev: number;
} {
  const railDev = Math.max(1, Math.round((lw * dpr) / 3));
  const gapDev = Math.max(1, Math.round((lw * dpr) / 3));
  return { railDev, gapDev, spanDev: 2 * railDev + gapDev };
}

/**
 * Paint a `double` border edge as device-pixel-aligned rail/gap/rail FILLS,
 * centred on the axis-aligned segment (x1,y1)-(x2,y2) — horizontal when
 * `y1 === y2`, else vertical. Computing both rails from a single rounded band
 * origin in device space (rather than two independently crisp-snapped strokes,
 * which can re-collapse the gap) keeps the bands whole device pixels at any
 * sz/scale/dpr. The caller must set `ctx.fillStyle` (and any save/restore).
 */
export function fillDoubleBorder(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  lw: number,
  dpr: number,
): void {
  const { railDev, gapDev, spanDev } = doubleRailGeometry(lw, dpr);
  if (y1 === y2) {
    // Horizontal edge: centre the band on its y, snapped to a whole device row.
    const startDev = Math.round(y1 * dpr - spanDev / 2);
    ctx.fillRect(x1, startDev / dpr, x2 - x1, railDev / dpr);
    ctx.fillRect(x1, (startDev + railDev + gapDev) / dpr, x2 - x1, railDev / dpr);
  } else {
    const startDev = Math.round(x1 * dpr - spanDev / 2);
    ctx.fillRect(startDev / dpr, y1, railDev / dpr, y2 - y1);
    ctx.fillRect((startDev + railDev + gapDev) / dpr, y1, railDev / dpr, y2 - y1);
  }
}
