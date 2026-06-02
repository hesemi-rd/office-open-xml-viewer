// Excel-style "nice" value-axis scaling. Pure math (no canvas), extracted so it
// can be unit-tested and reused independently of the chart renderer.

/** A round major-unit step that yields roughly `targetSteps` gridlines across
 *  `range` (1 / 2 / 5 × 10ⁿ — Excel's default ladder). */
export function niceStep(range: number, targetSteps = 5): number {
  if (range === 0) return 1;
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
}

/** Excel's automatic value-axis maximum: the smallest major-unit multiple that
 *  is >= dataMax — no extra headroom step. A series filling 97% of the axis
 *  stays at 97%, exactly as Excel renders it. (Overlap with sibling shapes is
 *  handled by honoring <c:plotArea><c:manualLayout> — ECMA-376 §21.2.2.32 —
 *  not by inflating the axis.) */
export function niceAxisMax(dataMax: number, step: number): number {
  if (dataMax <= 0) return step;
  return Math.ceil(dataMax / step) * step;
}

/** Axis minimum for data that dips below zero: the largest major-unit multiple
 *  <= dataMin, dropping one extra step when the data sits exactly on a
 *  gridline so the lowest point isn't flush against the axis. Non-negative data
 *  anchors the axis at 0. */
export function niceAxisMin(dataMin: number, step: number): number {
  if (dataMin >= 0) return 0;
  const ax = Math.floor(dataMin / step) * step;
  return Math.abs(ax - dataMin) < step * 1e-9 ? ax - step : ax;
}
