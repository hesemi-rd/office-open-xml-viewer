import { describe, it, expect } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { __test_renderInlineImage, type DecodedImage } from './renderer';
import type { LayoutImageSeg } from './line-layout';

/**
 * Regression test for #766 — the shared core `renderChart()` had no top-level
 * save/restore, so per-family renderers (e.g. pie-chart outer-ring labels)
 * left `ctx.textAlign='center'` / `ctx.textBaseline='middle'` set on return.
 * docx's inline-chart draw site (`renderInlineImage`, renderer.ts ~line 5807)
 * calls `renderChart()` bare — no wrapping save/restore of its own — so any
 * text drawn immediately after a chart segment on the same ctx (the next run
 * on the same line, or a table cell drawn later in the same paint pass) would
 * inherit the leaked center-alignment/mid-baseline. This was concretely
 * visible in sample-25: the "Other countries" table cell overlapped its
 * neighbor because it was drawn center-aligned instead of left-aligned.
 *
 * The mock ctx below implements a REAL save/restore stack (unlike a bare
 * no-op stub) so the fix in `renderChart` — a single top-level
 * `ctx.save(); try { … } finally { ctx.restore(); }` — is actually exercised.
 */

interface TextCall { text: string; x: number; y: number; align: string; baseline: string }

function stackfulMockCtx(): { ctx: CanvasRenderingContext2D; texts: TextCall[] } {
  const texts: TextCall[] = [];
  const defaults = {
    font: '10px sans-serif',
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  let state: Record<string, unknown> = { ...defaults };
  const stack: Record<string, unknown>[] = [];
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'save':
          return () => stack.push({ ...state });
        case 'restore':
          return () => { const s = stack.pop(); if (s) state = s; };
        case 'measureText':
          return (t: string) => ({ width: String(t).length * 6 });
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, texts };
}

function pieChartSeg(over: Partial<ChartModel> = {}): LayoutImageSeg {
  const chart: ChartModel = {
    chartType: 'pie',
    title: null,
    categories: ['Brazil', 'Vietnam', 'Other countries'],
    series: [{ name: 'Prod', color: null, values: [51500, 28500, 61000] }],
    showDataLabels: true,
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
  return {
    imagePath: '', mimeType: '', widthPt: 300, heightPt: 200,
    anchor: false, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
    chart, measuredWidth: 300,
  };
}

describe('docx inline chart draw site does not leak canvas state (#766)', () => {
  it('a fillText issued right after an inline chart segment is left-aligned/alphabetic-baseline, not the chart-internal center/middle', () => {
    const { ctx, texts } = stackfulMockCtx();
    const images = new Map<string, DecodedImage>();

    __test_renderInlineImage(ctx, pieChartSeg(), 10, 220, 1, images);
    // Simulate the very next draw call on the same ctx — e.g. a table cell's
    // text drawn later in the same paint pass (sample-25's "Other countries").
    ctx.fillText('Other countries', 10, 500);

    const after = texts[texts.length - 1];
    expect(after.text).toBe('Other countries');
    expect(after.align).toBe('start');
    expect(after.baseline).toBe('alphabetic');
  });

  it('ctx.textAlign/textBaseline are byte-identical before and after drawing an inline chart segment', () => {
    const { ctx } = stackfulMockCtx();
    const images = new Map<string, DecodedImage>();
    const before = { textAlign: ctx.textAlign, textBaseline: ctx.textBaseline, font: ctx.font, fillStyle: ctx.fillStyle };

    __test_renderInlineImage(ctx, pieChartSeg(), 10, 220, 1, images);

    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
    expect(ctx.font).toBe(before.font);
    expect(ctx.fillStyle).toBe(before.fillStyle);
  });
});
