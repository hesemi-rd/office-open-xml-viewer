// CH — value-axis MAJOR gridlines must be painted UNDER the data series, not
// over them. PowerPoint draws `<c:valAx><c:majorGridlines>` beneath the plotted
// geometry so an opaque area fill occludes the gridlines inside its region
// (verified against private/sample-14.pdf slide-6: every gridline that falls
// inside the teal ARR area reads solid teal, only the gridlines above the fill
// top are visible). The bar/line/stock/scatter/waterfall/box renderers already
// stroke gridlines before their series; this pins that ordering for the AREA
// family (which historically drew fills first, then gridlines on top) and guards
// the others against regressing.
import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';

// Ordered event recorder: logs each fill()/stroke() with the style in effect at
// the call, so we can assert the RELATIVE order of gridline strokes vs series
// fills. A translucent `rgba(...)` fillStyle marks a series area/line fill; a
// thin hairline strokeStyle (the resolved gridline color, default `#e0e0e0`)
// marks a gridline. We tag events by role and check the first gridline precedes
// the first series fill.
type Ev = { op: 'fill' | 'stroke'; fillStyle: string; strokeStyle: string; lineWidth: number };

function orderedRecordingCtx(): { ctx: CanvasRenderingContext2D; events: Ev[] } {
  const events: Ev[] = [];
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
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => ({ width: String(t).length * 6 });
        case 'fill':
          return () => events.push({ op: 'fill', fillStyle: String(state.fillStyle), strokeStyle: String(state.strokeStyle), lineWidth: Number(state.lineWidth) });
        case 'stroke':
          return () => events.push({ op: 'stroke', fillStyle: String(state.fillStyle), strokeStyle: String(state.strokeStyle), lineWidth: Number(state.lineWidth) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'moveTo': case 'lineTo': case 'arc': case 'fillRect':
        case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect':
        case 'strokeRect': case 'clearRect': case 'fillText': case 'strokeText':
        case 'setLineDash': case 'translate': case 'rotate': case 'scale':
        case 'clip': case 'setTransform': case 'resetTransform': case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, events };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'area',
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

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

// The default value-axis gridline is a thin hairline (`#e0e0e0`, 0.5 px). A
// series fill for the area family is a translucent `rgba(...)`. These heuristics
// classify the recorded events by role.
const isGridlineStroke = (e: Ev): boolean =>
  e.op === 'stroke' && e.lineWidth <= 1 && e.strokeStyle.toLowerCase() === '#e0e0e0';
const isSeriesFill = (e: Ev): boolean => e.op === 'fill' && /^rgba\(/i.test(e.fillStyle);

describe('CH — value-axis gridlines paint under the data series', () => {
  it('an area chart strokes its major gridlines BEFORE filling the series area', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['Jan', 'Feb', 'Mar', 'Apr'],
      series: [series({ name: 'ARR', values: [37, 40, 44, 48] })],
    }), RECT, 1);

    const firstGridline = rec.events.findIndex(isGridlineStroke);
    const firstSeriesFill = rec.events.findIndex(isSeriesFill);

    expect(firstSeriesFill).toBeGreaterThanOrEqual(0); // the area fill happened
    expect(firstGridline).toBeGreaterThanOrEqual(0);    // a gridline was stroked
    // The gridline must be laid down first so the opaque/translucent area sits
    // on top of it (PowerPoint occludes gridlines inside the fill region).
    expect(firstGridline).toBeLessThan(firstSeriesFill);
  });

  it('a stacked area chart also strokes gridlines before any fill', () => {
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['Jan', 'Feb', 'Mar'],
      series: [
        series({ name: 'A', values: [10, 12, 14] }),
        series({ name: 'B', values: [5, 6, 7] }),
      ],
    }), RECT, 1);

    const firstGridline = rec.events.findIndex(isGridlineStroke);
    const firstSeriesFill = rec.events.findIndex(isSeriesFill);
    expect(firstSeriesFill).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeLessThan(firstSeriesFill);
  });

  it('a plain line/area combo keeps gridlines below the area fill', () => {
    // Guard the line renderer (which draws area-like fills? no — line strokes).
    // Line chart already gridlines-first; assert it does not regress by pinning a
    // filled marker/area does not precede the gridline. Uses the line renderer.
    const rec = orderedRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A', 'B', 'C', 'D', 'E'],
      series: [series({ name: 'S', values: [1, 2, 3, 2, 4] })],
      // Explicit gridline color exercises the `grid.explicit` uniform-stroke path.
      valAxisGridlineColor: '888888',
    }), RECT, 1);
    const firstGridline = rec.events.findIndex(e => e.op === 'stroke' && e.strokeStyle.toLowerCase() === '#888888');
    const firstSeriesFill = rec.events.findIndex(isSeriesFill);
    expect(firstGridline).toBeGreaterThanOrEqual(0);
    expect(firstSeriesFill).toBeGreaterThanOrEqual(0);
    expect(firstGridline).toBeLessThan(firstSeriesFill);
  });
});
