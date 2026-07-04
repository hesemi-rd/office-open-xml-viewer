import { describe, it, expect } from 'vitest';
import {
  niceStep,
  niceAxisMax,
  niceAxisMin,
  valueAxisScale,
  axisFraction,
  logAxisScale,
  fitTrendline,
} from './axis-scale.js';

describe('niceStep', () => {
  it('picks 1/2/5 × 10ⁿ for ~5 gridlines', () => {
    expect(niceStep(100)).toBe(20);  // raw 20 → 2×10
    expect(niceStep(50)).toBe(10);   // raw 10 → 1×10
    expect(niceStep(7)).toBe(1);     // raw 1.4 → 1×1
    expect(niceStep(40)).toBe(10);   // raw 8 → 1×10 (8 ≥ 7.5 → 10)
  });
  it('zero range falls back to 1', () => {
    expect(niceStep(0)).toBe(1);
  });
});

describe('niceAxisMax (Excel headroom: first major unit above Ymax + range/20)', () => {
  it('rounds up past the ~5% headroom to the next major unit', () => {
    expect(niceAxisMax(41, 10)).toBe(50);        // 41 + 2.05 = 43.05 → 50
    expect(niceAxisMax(9715, 2000)).toBe(12000); // 9715 + 485.75 = 10200.75 → 12000
  });
  it('adds headroom even when data sits on a gridline (not flush against the top)', () => {
    expect(niceAxisMax(40, 10)).toBe(50);   // 40 + 2 = 42 → 50
    expect(niceAxisMax(100, 20)).toBe(120); // 100 + 5 = 105 → 120
  });
  it('uses dataMin for the range', () => {
    // range 100-(-100)=200, headroom 10 → 110 → step 50 → 150
    expect(niceAxisMax(100, 50, -100)).toBe(150);
  });
  it('non-positive max returns one step', () => {
    expect(niceAxisMax(0, 10)).toBe(10);
    expect(niceAxisMax(-5, 10)).toBe(10);
  });
});

describe('niceAxisMin', () => {
  it('non-negative data anchors at 0', () => {
    expect(niceAxisMin(15, 10)).toBe(0);
    expect(niceAxisMin(0, 10)).toBe(0);
  });
  it('negative data floors to a major-unit multiple', () => {
    expect(niceAxisMin(-15, 10)).toBe(-20);
  });
  it('data exactly on a gridline drops one extra step', () => {
    expect(niceAxisMin(-20, 10)).toBe(-30);
  });
});

describe('valueAxisScale (one niceStep drives min, max and gridline step)', () => {
  it('positive data anchored at 0 (bar/area/radar style)', () => {
    // step = niceStep(41-0) = niceStep(41) = 10; min = 0; max = niceAxisMax(41,10,0) = 50
    expect(valueAxisScale(0, 41)).toEqual({ min: 0, max: 50, step: 10 });
  });
  it('negative data floors the min and widens the max with the niced min', () => {
    // step = niceStep(100-(-15)) = niceStep(115) = 20;
    // min = niceAxisMin(-15,20) = -20; max = niceAxisMax(100,20,-20) = ceil((100+6)/20)*20 = 120
    expect(valueAxisScale(-15, 100)).toEqual({ min: -20, max: 120, step: 20 });
  });
  it('explicit min/max override the computed bounds (step still from data range)', () => {
    // step = niceStep(41-0) = 10; explicit min -5, max 60 win
    expect(valueAxisScale(0, 41, -5, 60)).toEqual({ min: -5, max: 60, step: 10 });
  });
  it('a null explicit bound falls back to the auto value', () => {
    expect(valueAxisScale(0, 41, null, 60)).toEqual({ min: 0, max: 60, step: 10 });
    expect(valueAxisScale(0, 41, -5, null)).toEqual({ min: -5, max: 50, step: 10 });
  });
  it('a longer value axis gets a finer major unit (Excel axis-length model)', () => {
    // Excel's auto major unit targets ~1 gridline per GRIDLINE_SPACING_PT (40),
    // so the SAME data range yields more, finer gridlines on a longer axis.
    // range 44: default target (5) → step 10 (0–50, 5 intervals); a 380pt axis
    // (target round(380/40)=10) → step 5 (0–50, 10 intervals) — matching the
    // horizontal-bar value axis (sample-14 slide-9) vs PowerPoint.
    expect(valueAxisScale(0, 44)).toEqual({ min: 0, max: 50, step: 10 });
    expect(valueAxisScale(0, 44, undefined, undefined, 380)).toEqual({ min: 0, max: 50, step: 5 });
  });
  it('data 3.5 → max 4 with step 0.5 (fine-grained positive range)', () => {
    // step = niceStep(3.5) = 0.5; min = 0; max = niceAxisMax(3.5,0.5,0):
    //   3.5 + 3.5/20 = 3.675 → ceil(3.675/0.5)*0.5 = 4
    expect(valueAxisScale(0, 3.5)).toEqual({ min: 0, max: 4, step: 0.5 });
  });
  it('data 0.1129 → max 0.12 with step 0.02 (sub-unit range)', () => {
    // step = niceStep(0.1129) = 0.02; min = 0;
    //   0.1129 + 0.1129/20 = 0.118545 → ceil(0.118545/0.02)*0.02 = 0.12
    const { min, max, step } = valueAxisScale(0, 0.1129);
    expect(min).toBe(0);
    expect(step).toBeCloseTo(0.02, 12);
    expect(max).toBeCloseTo(0.12, 12);
  });

  it('an explicit majorUnit overrides the auto step (min/max still auto)', () => {
    // <c:valAx><c:majorUnit val="25"/> forces the gridline spacing.
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, 25)).toEqual({
      min: 0,
      max: 50,
      step: 25,
    });
  });
  it('a null/undefined majorUnit keeps the auto step (byte-stable)', () => {
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, null)).toEqual(
      valueAxisScale(0, 41),
    );
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, undefined)).toEqual(
      valueAxisScale(0, 41),
    );
  });
  it('a non-positive majorUnit is ignored (no infinite gridline loop)', () => {
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, 0)).toEqual(
      valueAxisScale(0, 41),
    );
    expect(valueAxisScale(0, 41, undefined, undefined, undefined, -5)).toEqual(
      valueAxisScale(0, 41),
    );
  });
});

describe('axisFraction (value → 0..1 position along an axis)', () => {
  it('linear, normal orientation is exactly (v - min) / (max - min) — byte-stable', () => {
    expect(axisFraction(5, 0, 10)).toBe(0.5);
    expect(axisFraction(0, 0, 10)).toBe(0);
    expect(axisFraction(10, 0, 10)).toBe(1);
    expect(axisFraction(2, -10, 10)).toBe(0.6);
  });
  it('reversed orientation (maxMin) flips the fraction', () => {
    expect(axisFraction(5, 0, 10, { reversed: true })).toBe(0.5);
    expect(axisFraction(0, 0, 10, { reversed: true })).toBe(1);
    expect(axisFraction(10, 0, 10, { reversed: true })).toBe(0);
  });
  it('log axis maps in log space', () => {
    // base-10 axis 1..1000 → 10 sits at log10(10/1)/log10(1000/1) = 1/3
    expect(axisFraction(10, 1, 1000, { logBase: 10 })).toBeCloseTo(1 / 3, 12);
    expect(axisFraction(100, 1, 1000, { logBase: 10 })).toBeCloseTo(2 / 3, 12);
    expect(axisFraction(1, 1, 1000, { logBase: 10 })).toBe(0);
    expect(axisFraction(1000, 1, 1000, { logBase: 10 })).toBeCloseTo(1, 12);
  });
  it('log axis + reversed composes both', () => {
    expect(axisFraction(10, 1, 1000, { logBase: 10, reversed: true })).toBeCloseTo(2 / 3, 12);
  });
  it('degenerate zero range returns 0 (no NaN)', () => {
    expect(axisFraction(5, 5, 5)).toBe(0);
  });
});

describe('logAxisScale (power-of-base bounds + gridline exponents)', () => {
  it('snaps bounds down/up to powers of the base and lists the decade lines', () => {
    // data 3..700, base 10 → min 1 (10^0), max 1000 (10^3), lines 1,10,100,1000
    const s = logAxisScale(3, 700, 10);
    expect(s.min).toBe(1);
    expect(s.max).toBe(1000);
    expect(s.lines).toEqual([1, 10, 100, 1000]);
  });
  it('explicit min/max override the snapped bounds', () => {
    const s = logAxisScale(3, 700, 10, 1, 100);
    expect(s.min).toBe(1);
    expect(s.max).toBe(100);
    expect(s.lines).toEqual([1, 10, 100]);
  });
  it('clamps a non-positive data minimum up to the base (log undefined at <= 0)', () => {
    // data 0..500 can't take log(0); floor to the base's smallest positive decade.
    const s = logAxisScale(0, 500, 10);
    expect(s.min).toBeGreaterThan(0);
    expect(s.max).toBe(1000);
  });
});

describe('fitTrendline', () => {
  it('linear least squares recovers a perfect line', () => {
    // y = 2x + 1 at x = 0,1,2,3
    const t = fitTrendline([0, 1, 2, 3], [1, 3, 5, 7], 'linear');
    expect(t.xs).toEqual([0, 3]);
    expect(t.ys[0]).toBeCloseTo(1, 9);
    expect(t.ys[1]).toBeCloseTo(7, 9);
  });
  it('linear fit through noisy data has the least-squares slope', () => {
    // x 0..3, y 1,2,2,5 → slope = (nΣxy-ΣxΣy)/(nΣx²-(Σx)²) = (4*29-6*10)/(4*14-36)=56/20=1.2
    const t = fitTrendline([0, 1, 2, 3], [1, 2, 2, 5], 'linear');
    const slope = (t.ys[1] - t.ys[0]) / (t.xs[1] - t.xs[0]);
    expect(slope).toBeCloseTo(1.2, 9);
  });
  it('linear fit honors a forced intercept', () => {
    // Force b=0: m = Σxy/Σx². Σxy = 0·1+1·2+2·2+3·5 = 21; Σx² = 14 → m = 1.5.
    const t = fitTrendline([0, 1, 2, 3], [1, 2, 2, 5], 'linear', { intercept: 0 });
    expect(t.ys[0]).toBeCloseTo(0, 9); // line passes through (0,0)
    const slope = (t.ys[1] - t.ys[0]) / (t.xs[1] - t.xs[0]);
    expect(slope).toBeCloseTo(21 / 14, 9);
  });
  it('moving average (period 2) trails the mean of the last two points', () => {
    const t = fitTrendline([0, 1, 2, 3], [10, 20, 30, 40], 'movingAvg', { period: 2 });
    expect(t.xs).toEqual([1, 2, 3]);
    expect(t.ys).toEqual([15, 25, 35]);
  });
  it('unsupported types return empty (parse-only for now)', () => {
    expect(fitTrendline([0, 1, 2], [1, 2, 4], 'exp')).toEqual({ xs: [], ys: [] });
    expect(fitTrendline([0, 1, 2], [1, 2, 4], 'poly')).toEqual({ xs: [], ys: [] });
  });
  it('too few points returns empty', () => {
    expect(fitTrendline([0], [1], 'linear')).toEqual({ xs: [], ys: [] });
  });
});
