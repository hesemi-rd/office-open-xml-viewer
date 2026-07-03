// Behavioral tests for the chart-correctness fixes:
//   CH1 — bar/column negative values extend downward from the zero line
//   CH2 — stackedLine / stackedLinePct series are cumulatively stacked
//   CH3 — tick / data labels use the locale-independent §18.8.30 formatter
//
// These assert observable geometry (fillRect bounds, gridline label text)
// captured through a lightweight recording context, complementing the
// draw-call-signature characterization test.

import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';

interface RectCall { x: number; y: number; w: number; h: number; fs: string }
interface TextCall { text: string; x: number; y: number }

interface Recorded {
  ctx: CanvasRenderingContext2D;
  rects: RectCall[];
  texts: TextCall[];
}

/** Minimal recording 2D context: captures fillRect + fillText, tracks the
 *  handful of state props the renderer reads, and models text width. */
function recordingCtx(): Recorded {
  const rects: RectCall[] = [];
  const texts: TextCall[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) =>
            rects.push({ x, y, w, h, fs: String(state.fillStyle) });
        case 'fillText':
          return (text: string, x: number, y: number) => texts.push({ text, x, y });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'fill': case 'stroke': case 'moveTo': case 'lineTo': case 'arc':
        case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect':
        case 'strokeRect': case 'clearRect': case 'strokeText': case 'setLineDash':
        case 'translate': case 'rotate': case 'scale': case 'clip':
        case 'setTransform': case 'resetTransform': case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, rects, texts };
}

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'clusteredBar',
    title: null,
    categories: [],
    series: [],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: false,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'out',
    catAxisMajorTickMark: 'out',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: null,
    valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    ...over,
  };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

describe('CH1 — negative bar/column values extend from the zero line', () => {
  it('a column chart draws negative bars below the zero line and positive bars above', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, -10] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    // Two bars, one per category.
    expect(bars.length).toBe(2);
    const [pos, neg] = bars;
    // Symmetric data (+10 / -10) → the zero line sits mid-plot and the two bars
    // have equal height. The positive bar's bottom edge equals the negative
    // bar's top edge: they meet at the shared zero line.
    const posBottom = pos.y + pos.h;
    const negTop = neg.y;
    expect(negTop).toBeCloseTo(posBottom, 4); // shared zero line
    // Negative bar hangs BELOW the zero line, positive bar sits ABOVE it.
    expect(neg.y).toBeGreaterThan(pos.y);
    expect(neg.h).toBeGreaterThan(0);
    // Equal magnitudes → equal bar heights.
    expect(neg.h).toBeCloseTo(pos.h, 4);
  });

  it('the value axis includes negative tick labels when data dips below zero', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({ name: 'S', values: [-40] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels.some(l => l.startsWith('-'))).toBe(true);
  });

  it('a horizontal bar chart draws negative bars left of the zero line', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, -10] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [pos, neg] = bars;
    // Positive bar starts at the zero line and extends right; negative bar ends
    // at the zero line and extends left, so its right edge equals the positive
    // bar's left edge.
    expect(neg.x + neg.w).toBeCloseTo(pos.x, 4);
    expect(neg.x).toBeLessThan(pos.x);
    expect(neg.w).toBeCloseTo(pos.w, 4);
  });

  it('positive-only data keeps the axis anchored at 0 (pre-fix behavior)', () => {
    // Regression guard: min degenerates to 0 so nothing about a positive-only
    // chart changes. Zero-line bottom edge == plot bottom.
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 20] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    // All bars share the same bottom edge (the axis at 0), none extend below it.
    const bottoms = bars.map(b => b.y + b.h);
    expect(bottoms[0]).toBeCloseTo(bottoms[1], 4);
    // No negative tick labels.
    expect(rec.texts.every(t => !t.text.startsWith('-'))).toBe(true);
  });

  it('stacked columns accumulate positives up and negatives down separately', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'stackedBar',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),
        series({ name: 'N', values: [-20] }),
      ],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar sits above the zero line; negative bar below. They meet at
    // the zero line (positive bottom == negative top).
    expect(nBar.y).toBeCloseTo(p.y + p.h, 4);
    expect(nBar.h).toBeGreaterThan(0);
  });
});

describe('CH6 — negative bar data-label placement mirrors the positive convention (§21.2.2.16)', () => {
  // Coverage for drawBarDataLabel's `negative` branch. A single chart holds two
  // categories with a symmetric +37 / -37 value, so BOTH bars share one plot and
  // one axis (a symmetric ±37 range) — the geometry is a clean mirror across the
  // zero line. For each dLblPos the negative label must land on the mirror side
  // of the positive label relative to that shared zero line. "37" / "-37" are
  // not round gridline values, so each data-label text is unambiguous among the
  // recorded fillText calls, and each bar is matched to its label by the shared
  // cross-axis center.
  function renderMirrorBars(
    chartType: 'clusteredBar' | 'clusteredBarH',
    dataLabelPosition: string,
  ): Recorded {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['P', 'N'],
      series: [series({ name: 'S', values: [37, -37] })],
      showDataLabels: true,
      dataLabelPosition,
    }), RECT, 1);
    return rec;
  }
  const labelPos = (rec: Recorded, text: string): TextCall => {
    const hit = rec.texts.find(t => t.text === text);
    expect(hit, `data label "${text}" was drawn`).toBeDefined();
    return hit as TextCall;
  };
  // Match each value bar to its label by the cross-axis center they share
  // (x-center for columns, y-center for horizontal bars).
  const barFor = (rec: Recorded, lbl: TextCall, axis: 'v' | 'h'): RectCall => {
    const center = (b: RectCall) => axis === 'v' ? b.x + b.w / 2 : b.y + b.h / 2;
    const key = axis === 'v' ? lbl.x : lbl.y;
    let best: RectCall | undefined;
    let bestD = Infinity;
    for (const b of rec.rects) {
      const d = Math.abs(center(b) - key);
      if (d < bestD) { bestD = d; best = b; }
    }
    expect(best).toBeDefined();
    return best as RectCall;
  };

  describe('vertical columns', () => {
    for (const pos of ['outEnd', 'inEnd', 'inBase', 'ctr']) {
      it(`${pos}: the negative label mirrors the positive label across the zero line`, () => {
        const rec = renderMirrorBars('clusteredBar', pos);
        const posLbl = labelPos(rec, '37');
        const negLbl = labelPos(rec, '-37');
        const posBar = barFor(rec, posLbl, 'v');   // sits ABOVE the zero line
        const negBar = barFor(rec, negLbl, 'v');   // hangs BELOW the zero line
        // Each label is horizontally centered on its own bar.
        expect(posLbl.x).toBeCloseTo(posBar.x + posBar.w / 2, 4);
        expect(negLbl.x).toBeCloseTo(negBar.x + negBar.w / 2, 4);
        // Symmetric ±37 → equal bar heights, bars meeting at the shared zero line.
        expect(negBar.h).toBeCloseTo(posBar.h, 3);
        const zeroLine = posBar.y + posBar.h;            // positive bottom == neg top
        expect(negBar.y).toBeCloseTo(zeroLine, 3);
        // The positive bar's value edge is its TOP; the negative's is its BOTTOM.
        const posValueEdge = posBar.y;                   // top edge
        const negValueEdge = negBar.y + negBar.h;        // bottom edge
        if (pos === 'ctr') {
          expect(posLbl.y).toBeCloseTo(posBar.y + posBar.h / 2, 4);
          expect(negLbl.y).toBeCloseTo(negBar.y + negBar.h / 2, 4);
          // The two centers are mirror images across the zero line.
          expect(negLbl.y - zeroLine).toBeCloseTo(zeroLine - posLbl.y, 3);
        } else if (pos === 'outEnd' || pos === 'inEnd') {
          // Positive label offset from its top edge mirrors the negative label
          // offset from its bottom edge (positive sits above → −, negative below → +).
          const posOff = posLbl.y - posValueEdge;
          const negOff = negLbl.y - negValueEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        } else {
          // inBase: anchored at the zero-line (base) edge for both signs.
          const posBaseEdge = posBar.y + posBar.h;       // bottom (zero line)
          const negBaseEdge = negBar.y;                  // top (zero line)
          const posOff = posLbl.y - posBaseEdge;
          const negOff = negLbl.y - negBaseEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        }
      });
    }
  });

  describe('horizontal bars', () => {
    for (const pos of ['outEnd', 'inEnd', 'inBase', 'ctr']) {
      it(`${pos}: the negative label mirrors the positive label across the zero line`, () => {
        const rec = renderMirrorBars('clusteredBarH', pos);
        const posLbl = labelPos(rec, '37');
        const negLbl = labelPos(rec, '-37');
        const posBar = barFor(rec, posLbl, 'h');   // extends RIGHT of the zero line
        const negBar = barFor(rec, negLbl, 'h');   // extends LEFT of the zero line
        // Each label is vertically centered on its own bar. The recorded rect is
        // fillRect(bx, by, barL, barW), so its HEIGHT is the bar thickness.
        expect(posLbl.y).toBeCloseTo(posBar.y + posBar.h / 2, 4);
        expect(negLbl.y).toBeCloseTo(negBar.y + negBar.h / 2, 4);
        // Symmetric ±37 → equal bar lengths, meeting at the shared zero line.
        expect(negBar.w).toBeCloseTo(posBar.w, 3);
        const zeroLine = posBar.x;                        // positive left == neg right
        expect(negBar.x + negBar.w).toBeCloseTo(zeroLine, 3);
        if (pos === 'ctr') {
          expect(posLbl.x).toBeCloseTo(posBar.x + posBar.w / 2, 4);
          expect(negLbl.x).toBeCloseTo(negBar.x + negBar.w / 2, 4);
          expect(negLbl.x - zeroLine).toBeCloseTo(zeroLine - posLbl.x, 3);
        } else if (pos === 'outEnd' || pos === 'inEnd') {
          // Positive value edge is the RIGHT edge; negative value edge the LEFT.
          const posValueEdge = posBar.x + posBar.w;
          const negValueEdge = negBar.x;
          const posOff = posLbl.x - posValueEdge;
          const negOff = negLbl.x - negValueEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        } else {
          // inBase: zero-line edge. Positive base is the LEFT edge, negative base
          // the RIGHT edge — mirrored across the zero line.
          const posBaseEdge = posBar.x;                  // left (zero line)
          const negBaseEdge = negBar.x + negBar.w;       // right (zero line)
          const posOff = posLbl.x - posBaseEdge;
          const negOff = negLbl.x - negBaseEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        }
      });
    }
  });
});

describe('CH7 — percentStacked normalizes signed values against per-category Σ|v| (§21.2.2.76)', () => {
  // Positive contributions stack up/right, negatives down/left; each series is
  // normalized to (v / Σ|v|)·100 so the axis spans −100..100.
  it('vertical percentStacked: positives stack above zero, negatives below, normalized to Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),   // +30
        series({ name: 'N', values: [-10] }),  // -10  → Σ|v| = 40
      ],
    }), RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar sits above the zero line, negative bar below; they meet at it.
    expect(nBar.y).toBeCloseTo(p.y + p.h, 3);          // shared zero line
    expect(nBar.y).toBeGreaterThan(p.y);               // negative is lower
    // Normalized magnitudes: +30/40 = 75% up, -10/40 = 25% down. Same axis
    // scale (px per percent) → the positive bar is 3× the negative bar's height.
    expect(p.h / nBar.h).toBeCloseTo(3, 2);
    // The value axis carries the ±100 percentStacked gridlines (plus headroom,
    // so the outermost ticks sit at ±120, matching the line/area pct convention).
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    expect(nums).toContain(100);
    expect(nums).toContain(-100);
    expect(Math.min(...nums)).toBeLessThanOrEqual(-100);
    expect(Math.min(...nums)).toBeGreaterThanOrEqual(-120);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('horizontal percentStacked: positives stack right, negatives left, normalized to Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarHPct',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),   // +30 → right
        series({ name: 'N', values: [-10] }),  // -10 → left, Σ|v| = 40
      ],
    }), RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar extends right of the zero line, negative left; they meet at it.
    expect(nBar.x + nBar.w).toBeCloseTo(p.x, 3);       // shared zero line
    expect(nBar.x).toBeLessThan(p.x);                  // negative is to the left
    // +30/40 = 75% right vs -10/40 = 25% left → 3× the width.
    expect(p.w / nBar.w).toBeCloseTo(3, 2);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    expect(nums).toContain(100);
    expect(nums).toContain(-100);
    expect(Math.min(...nums)).toBeLessThanOrEqual(-100);
    expect(Math.min(...nums)).toBeGreaterThanOrEqual(-120);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('multi-category percentStacked: each category normalizes to its own Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'P', values: [10, 40] }),  // A: Σ|v|=20  B: Σ|v|=50
        series({ name: 'N', values: [-10, -10] }),
      ],
    }), RECT, 1);
    const bars = rec.rects;
    // Two categories × two series = four bars, in draw order: A/P, A/N, B/P, B/N.
    expect(bars.length).toBe(4);
    const [aP, aN, bP, bN] = bars;
    // Category A: 10 and -10 of Σ|v|=20 → 50% up, 50% down → equal heights.
    expect(aP.h).toBeCloseTo(aN.h, 2);
    // Category B: 40 and -10 of Σ|v|=50 → 80% up, 20% down → positive is 4× taller.
    expect(bP.h / bN.h).toBeCloseTo(4, 2);
    // Per-category normalization (not a global Σ): A's +50% bar and B's +80% bar
    // are NOT the same height even though A/P is the larger raw share of A.
    expect(bP.h).toBeGreaterThan(aP.h);
  });
});

describe('CH2 — stackedLine / stackedLinePct stack cumulatively', () => {
  it('stackedLine plots the second series at the cumulative sum', () => {
    // Two flat series (all 10, all 20). Stacked, the second line rides at
    // y=30 across every category; unstacked it would ride at y=20. We detect
    // stacking by the axis maximum: a cumulative 30 forces a taller axis than
    // an un-stacked max of 20 would.
    const stackedRec = recordingCtx();
    renderChart(stackedRec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'S1', values: [10, 10, 10] }),
        series({ name: 'S2', values: [20, 20, 20] }),
      ],
    }), RECT, 1);

    const plainRec = recordingCtx();
    renderChart(plainRec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'S1', values: [10, 10, 10] }),
        series({ name: 'S2', values: [20, 20, 20] }),
      ],
    }), RECT, 1);

    const stackedTop = Math.max(...stackedRec.texts
      .map(t => Number(t.text)).filter(v => Number.isFinite(v)));
    const plainTop = Math.max(...plainRec.texts
      .map(t => Number(t.text)).filter(v => Number.isFinite(v)));
    // Stacking pushes the cumulative maximum (30) above the plain per-series
    // maximum (20), so the auto axis top must be strictly higher.
    expect(stackedTop).toBeGreaterThan(plainTop);
  });

  it('stackedLinePct normalizes each category to 100%', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLinePct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // The cumulative top series always reaches exactly 100% per category, so the
    // axis carries a 100 gridline. Raw magnitudes (max cumulative 40) never
    // appear — the axis is normalized, not driven by the raw sums.
    expect(nums).toContain(100);
    // ...and the axis top is a round value just above 100 (headroom), never the
    // raw cumulative magnitude of 40.
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('plain line is unaffected (per-series max drives the axis)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 10] }),
        series({ name: 'S2', values: [20, 20] }),
      ],
    }), RECT, 1);
    const top = Math.max(...rec.texts.map(t => Number(t.text)).filter(Number.isFinite));
    // Un-stacked: axis reflects the single-series max (20) plus headroom, not 30.
    expect(top).toBeLessThan(30);
  });
});

describe('CH4 — stackedAreaPct normalizes like the line/bar percentStacked convention', () => {
  it('stackedAreaPct normalizes each category to 100%', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedAreaPct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // The cumulative top series always reaches exactly 100% per category, so the
    // axis carries a 100 gridline. Raw magnitudes (max cumulative 40) never
    // appear — the axis is normalized, not driven by the raw sums (this was Red
    // before the fix: stackedAreaPct was treated identically to stackedArea, so
    // the axis topped out at the raw cumulative 40 instead of 100).
    expect(nums).toContain(100);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('stackedArea (non-percent) is unaffected — axis reflects the raw cumulative sum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // Raw cumulative max per category is 40 (10+30 / 30+10); the axis must scale
    // to that magnitude, not be normalized to 100.
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(40);
    expect(nums).not.toContain(100);
  });
});

describe('CH5 — category axis numFmt applies to category tick labels (§21.2.2.71)', () => {
  // dateAx / numeric category axes cache the categories as Excel serial numbers
  // ("44927"). Before the fix the renderer drew those raw serials; now the
  // catAxisFormatCode is applied so a time-series line/bar shows real dates.
  const DATE_CATS = ['44927', '44958', '44986']; // 2023-01-01 / 02-01 / 03-01

  it('a line chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels).toContain('2/1/2023');
    expect(labels).toContain('3/1/2023');
    // The raw serials must NOT appear as category labels anymore.
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('a column chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('a horizontal bar chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('an area chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('string categories stay verbatim even when a format code is present', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Q1', 'Q2', 'Q3'],
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('Q1');
    expect(labels).toContain('Q2');
    expect(labels).toContain('Q3');
  });

  it('numeric categories with no format code render as raw text (unchanged)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: DATE_CATS,
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('44927');
    expect(labels.some(l => l === '1/1/2023')).toBe(false);
  });
});

describe('CH3 — labels are locale-independent (§18.8.30)', () => {
  // `toLocaleString()` groups thousands in every common locale, so an explicit
  // no-separator format code ("0") is the discriminator: the §18.8.30 engine
  // honors it (no commas), while toLocaleString ignores it and always inserts
  // the host locale's group separator. The old code called toLocaleString and
  // dropped the format code entirely, so these tests were Red before the fix.
  it('waterfall data labels honor the format code (no host-locale grouping)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'End'],
      series: [series({ name: 'W', values: [1234567, 0] })],
      subtotalIndices: [1],
      dataLabelFormatCode: '0',
    }), RECT, 1);
    // The 1234567 subtotal bar's label must be un-grouped ("1234567"), proving
    // it went through the §18.8.30 engine rather than toLocaleString().
    expect(rec.texts.some(t => t.text.includes('1234567'))).toBe(true);
    expect(rec.texts.every(t => !t.text.includes('1,234,567'))).toBe(true);
  });

  it('waterfall negative data labels keep the △ marker and honor the format code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Drop', 'End'],
      series: [series({ name: 'W', values: [2000000, -1234567, 765433] })],
      subtotalIndices: [2],
      dataLabelFormatCode: '0',
    }), RECT, 1);
    // Negative bar: △ prefix + un-grouped magnitude from the engine.
    expect(rec.texts.some(t => t.text.includes('△') && t.text.includes('1234567'))).toBe(true);
    expect(rec.texts.every(t => !t.text.includes('1,234,567'))).toBe(true);
  });

  it('waterfall value-axis labels honor the format code (through the §18.8.30 engine)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'End'],
      series: [series({ name: 'W', values: [1000000, 0] })],
      subtotalIndices: [1],
      valAxisFormatCode: '0',
    }), RECT, 1);
    // A no-separator format code must suppress grouping. The old code ignored
    // valAxisFormatCode and always grouped via toLocaleString(), so a "1,000,000"
    // tick label would appear — after the fix the ticks are un-grouped.
    expect(rec.texts.every(t => !t.text.includes('1,000,000'))).toBe(true);
    expect(rec.texts.some(t => /^\d{4,}$/.test(t.text))).toBe(true);
  });
});
