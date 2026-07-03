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
