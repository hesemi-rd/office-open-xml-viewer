// Shared chart axis-style helpers. The cartesian value-axis renderers
// (bar / line / area / scatter) and the combo secondary axis all resolved the
// same `<c:*Ax><c:spPr><a:ln>` colour/width and the same crossBetween default
// inline; these extract that logic so it lives in one place.

import type { ChartModel } from '../types/chart';
import { EMU_PER_PT } from '../units.js';

/** An axis rule's `<a:ln@w>` (EMU) → canvas px at the current display scale.
 *  `ptToPx` is px-per-point. Floored at 0.5 px so a thin rule stays visible;
 *  an absent width falls back to a 1 px hairline. */
export function axisLineWidthPx(widthEmu: number | null | undefined, ptToPx: number): number {
  return widthEmu ? Math.max(0.5, widthEmu / EMU_PER_PT) * ptToPx : 1;
}

/** Resolve an axis rule's `{ color, width }` from its parsed `<a:ln>` parts.
 *  Colour defaults to Office's faint default rule (`#aaa`) and width to 1 px.
 *  Used for the category/value axes (`chart.*AxisLineColor/WidthEmu`) and the
 *  secondary value axis (`sec.lineColor/lineWidthEmu`) — pass whichever fields
 *  apply. (Scatter keeps its own `#888`/`undefined` defaults, so it uses
 *  `axisLineWidthPx` directly rather than this.) */
export function resolveAxisLine(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  ptToPx: number,
): { color: string; width: number } {
  return {
    color: color ? `#${color}` : '#aaa',
    width: axisLineWidthPx(widthEmu, ptToPx),
  };
}

/** A resolved value-axis MAJOR gridline stroke: `{ color, width }` from the
 *  parsed `<c:majorGridlines><c:spPr><a:ln>` parts (`chart.valAxisGridlineColor`
 *  / `valAxisGridlineWidthEmu`).
 *
 *  When the file supplies neither, this reproduces the renderer's historical
 *  faint hairline — `#e0e0e0` at 0.5 px — so charts without a styled gridline
 *  stay byte-stable. An explicit `<a:solidFill>` color (e.g. `accent3` →
 *  `#8FA878`) or `<a:ln w>` width overrides that default; the width goes through
 *  {@link axisLineWidthPx}, which floors at 0.5 px so a thin (0.25 pt) rule
 *  stays visible on a device pixel grid.
 *
 *  Callers still special-case the value-0 gridline (`#aaa`/1 px emphasis) — but
 *  ONLY when there is no explicit gridline color. When the file pins a gridline
 *  color, PowerPoint strokes every major gridline in that one color uniformly
 *  (there is no separate baseline treatment), so callers pass `hasExplicitColor`
 *  to suppress the zero-line override. */
export function resolveGridline(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  ptToPx: number,
): { color: string; width: number } {
  return {
    color: color ? `#${color}` : '#e0e0e0',
    // Absent width keeps the historical 0.5 px hairline (not axisLineWidthPx's
    // 1 px axis-rule fallback) so an unstyled gridline is byte-identical.
    width: widthEmu ? axisLineWidthPx(widthEmu, ptToPx) : 0.5,
  };
}

/** Category-axis crossBetween: by default categories occupy a band and points
 *  sit at the band centre ("between"); `"midCat"` anchors them on the dividers.
 *  ECMA-376 §21.2.2.32 leaves the default application-defined — Office (and we)
 *  use "between". */
export function isCrossBetween(chart: Pick<ChartModel, 'catAxisCrossBetween'>): boolean {
  return chart.catAxisCrossBetween !== 'midCat';
}
