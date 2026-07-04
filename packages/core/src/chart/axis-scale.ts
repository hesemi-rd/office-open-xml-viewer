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

/** Excel / PowerPoint automatic value-axis maximum. Microsoft's documented
 *  algorithm (per Peltier Tech) is "the first major unit above
 *  `Ymax + (Ymax − Ymin)/20`": ~5% of the data range is added as headroom so the
 *  tallest series sits just below the top gridline rather than flush against it,
 *  then the result is rounded up to the next major unit. `dataMin` is the axis
 *  minimum (0 for bar/column charts; the data minimum otherwise).
 *
 *  The major unit itself is Excel-proprietary (it varies with plot size, tick
 *  font, etc. and is not documented), so we approximate it with `niceStep`; the
 *  computed max can therefore differ from PowerPoint by one major unit on some
 *  charts. */
export function niceAxisMax(dataMax: number, step: number, dataMin = 0): number {
  if (dataMax <= 0) return step;
  const withHeadroom = dataMax + (dataMax - dataMin) / 20;
  return Math.ceil(withHeadroom / step) * step;
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

/** Excel-style auto value-axis bounds + major unit. ONE `niceStep` (of the data
 *  range) drives the rounded min, max AND the gridline step, so they can never
 *  desync. Explicit `<c:valAx><c:scaling><c:min/max>` wins. The auto major unit
 *  is Excel-proprietary (not in ECMA-376); niceStep approximates it. */
/** Target gridline spacing in POINTS. Excel's auto major unit is not a fixed
 *  gridline count — it targets a roughly constant on-screen spacing, so a long
 *  axis (e.g. a horizontal bar chart's wide value axis) gets MORE, finer
 *  gridlines than a short one of the same data range. Empirically ~one major
 *  gridline per this many points reproduces PowerPoint across sample-14's
 *  column / area / horizontal-bar / secondary-axis charts. This is a runtime
 *  Excel behavior (not in ECMA-376); the constant is the one tunable. */
const GRIDLINE_SPACING_PT = 40;

/** Pick the `niceStep` target-gridline count for an axis of `axisLenPt` points.
 *  Falls back to 5 (the legacy fixed target) when the length is unknown. */
function targetStepsForAxis(axisLenPt?: number): number {
  if (axisLenPt == null || !isFinite(axisLenPt) || axisLenPt <= 0) return 5;
  return Math.min(15, Math.max(3, Math.round(axisLenPt / GRIDLINE_SPACING_PT)));
}

export function valueAxisScale(
  dataMin: number, dataMax: number,
  explicitMin?: number | null, explicitMax?: number | null,
  axisLenPt?: number,
  majorUnit?: number | null,
): { min: number; max: number; step: number } {
  // A file-specified `<c:valAx><c:majorUnit>` (ECMA-376 §21.2.2.103) overrides
  // the auto "nice" step. Non-positive / non-finite values are ignored (they'd
  // wedge the gridline loop) so an absent/invalid majorUnit is byte-stable.
  const autoStep = niceStep(dataMax - dataMin, targetStepsForAxis(axisLenPt));
  const step = majorUnit != null && isFinite(majorUnit) && majorUnit > 0 ? majorUnit : autoStep;
  // The auto min/max still round against the AUTO step (Excel derives the bounds
  // before the user's manual major unit is applied); explicit scaling wins.
  const min = explicitMin ?? niceAxisMin(dataMin, autoStep);
  const max = explicitMax ?? niceAxisMax(dataMax, autoStep, min);
  return { min, max, step };
}

/** Options for {@link axisFraction}: a logarithmic base and/or a reversed
 *  (`maxMin`) orientation. Both default off, in which case the fraction is the
 *  plain linear `(v - min) / (max - min)` — byte-identical to the renderers'
 *  historical inline math. */
export interface AxisFractionOpts {
  /** `<c:scaling><c:logBase val>` (ECMA-376 §21.2.2.98). When set (>= 2) the
   *  value maps in log space. min/max must be positive. */
  logBase?: number | null;
  /** `<c:scaling><c:orientation val="maxMin">` — reverse the axis. */
  reversed?: boolean;
}

/** Map a value to its 0..1 position along an axis spanning `min`..`max`.
 *
 *  This is the single shared primitive the per-chart-type renderers build their
 *  `toY` / `valX` closures on: a caller does `py0 + ph - axisFraction(v, min,
 *  max) * ph`. With no options it returns exactly `(v - min) / (max - min)`, so
 *  a linear, normally-oriented axis is byte-stable. A log base maps in log space
 *  (gridlines fall on powers of the base); `reversed` flips the fraction for a
 *  `maxMin` orientation. A degenerate zero range yields 0 (no NaN/∞). */
export function axisFraction(
  v: number, min: number, max: number, opts?: AxisFractionOpts,
): number {
  let frac: number;
  const logBase = opts?.logBase;
  if (logBase != null && isFinite(logBase) && logBase >= 2 && min > 0 && max > 0) {
    const lo = Math.log(min);
    const hi = Math.log(max);
    const denom = hi - lo;
    frac = denom === 0 ? 0 : (Math.log(Math.max(v, Number.MIN_VALUE)) - lo) / denom;
  } else {
    const denom = max - min;
    frac = denom === 0 ? 0 : (v - min) / denom;
  }
  return opts?.reversed ? 1 - frac : frac;
}

/** Logarithmic value-axis bounds + gridline decades. Snaps `dataMin` down and
 *  `dataMax` up to whole powers of `base` (Excel's log-axis behavior) and lists
 *  every power-of-base gridline in `[min, max]`. Explicit `<c:scaling><c:min /
 *  max>` override the snapped bounds. A non-positive `dataMin` (log undefined at
 *  <= 0) floors to the base's lowest decade at or below the smallest positive
 *  datum, defaulting to `base^0 = 1` when there is none. */
export function logAxisScale(
  dataMin: number, dataMax: number, base: number,
  explicitMin?: number | null, explicitMax?: number | null,
): { min: number; max: number; lines: number[] } {
  const b = isFinite(base) && base >= 2 ? base : 10;
  const logB = (x: number): number => Math.log(x) / Math.log(b);
  const posMax = dataMax > 0 ? dataMax : 1;
  // Floor the min to a decade; clamp a non-positive min up to the smallest
  // decade that is still <= the positive max.
  const rawMin = dataMin > 0 ? dataMin : posMax;
  const minExp = Math.floor(logB(rawMin));
  const maxExp = Math.ceil(logB(posMax));
  const min = explicitMin != null ? explicitMin : Math.pow(b, minExp);
  const max = explicitMax != null ? explicitMax : Math.pow(b, Math.max(maxExp, minExp + 1));
  const lines: number[] = [];
  const startExp = Math.ceil(logB(min) - 1e-9);
  const endExp = Math.floor(logB(max) + 1e-9);
  for (let e = startExp; e <= endExp; e++) lines.push(Math.pow(b, e));
  return { min, max, lines };
}
