import { describe, it, expect } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';

// ─────────────────────────────────────────────────────────────────────────────
// Oracle for C10 — pptx chart nested emit.
//
// Before this refactor the pptx renderer built the core `ChartModel` from a
// flat `ChartElement` via a 60-field object literal (the "adapter"). Now the
// Rust parser emits the `ChartModel` shape directly as `el.chart`, and the
// renderer passes it straight to `renderChart`.
//
// This test freezes that former literal as `legacyPptxChartModel()` (an exact
// copy of the deleted renderer code, kept here as an oracle) and proves the new
// pass-through is render-equivalent: given the SAME parsed source, the object
// the renderer now hands to `renderChart` (built here with the field-name map
// the Rust parser applies) deep-equals the object the old literal would have
// produced.
//
// `null` vs. missing key is normalized away (`stripNullish`) because the core
// renderer reads every optional field as `?? default` / `!= null` — an absent
// key and an explicit `null` render identically. Array elements (e.g. a `null`
// data point inside `values`) are preserved: only object properties are pruned.
// ─────────────────────────────────────────────────────────────────────────────

/** The former pptx flat chart element (pre-refactor shape). */
interface LegacyPptxChartElement {
  chartType: string;
  title: string | null;
  categories: string[];
  series: ChartModel['series'];
  showDataLabels: boolean;
  valMin: number | null;
  valMax: number | null;
  catAxisTitle?: string | null;
  valAxisTitle?: string | null;
  catAxisTitleSize?: number | null;
  catAxisTitleBold?: boolean | null;
  catAxisTitleColor?: string | null;
  valAxisTitleSize?: number | null;
  valAxisTitleBold?: boolean | null;
  valAxisTitleColor?: string | null;
  catAxisHidden: boolean;
  valAxisHidden: boolean;
  catAxisLineHidden?: boolean;
  valAxisLineHidden?: boolean;
  catAxisFontColor?: string | null;
  valAxisFontColor?: string | null;
  catAxisLineColor?: string | null;
  catAxisLineWidthEmu?: number | null;
  valAxisLineColor?: string | null;
  valAxisLineWidthEmu?: number | null;
  plotAreaBg: string | null;
  chartBg: string | null;
  showLegend: boolean;
  legendPos?: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  catAxisCrossBetween: string;
  valAxisMajorTickMark: string;
  catAxisMajorTickMark: string;
  titleFontSizeHpt: number | null;
  titleFontColor?: string | null;
  titleFontFace?: string | null;
  titleFontBold?: boolean | null;
  catAxisFontSizeHpt: number | null;
  valAxisFontSizeHpt: number | null;
  catAxisFontBold?: boolean | null;
  valAxisFontBold?: boolean | null;
  dataLabelFontSizeHpt: number | null;
  subtotalIndices: number[];
  barGapWidth?: number | null;
  barOverlap?: number | null;
  dataLabelPosition?: string | null;
  dataLabelFontColor?: string | null;
  dataLabelFormatCode?: string | null;
  valAxisFormatCode?: string | null;
  plotAreaManualLayout?: ChartModel['plotAreaManualLayout'];
  scatterStyle?: string | null;
  radarStyle?: string | null;
  chartBorderColor?: string | null;
  chartBorderWidthEmu?: number | null;
  secondaryValAxis?: ChartModel['secondaryValAxis'];
}

/**
 * FROZEN copy of the deleted renderer literal (pptx renderer.ts, pre-C10). Do
 * not "modernize" this — it is the oracle; its output is the ground truth the
 * new pass-through must match.
 */
function legacyPptxChartModel(el: LegacyPptxChartElement): ChartModel {
  return {
    chartType: el.chartType,
    title: el.title,
    categories: el.categories,
    series: el.series,
    showDataLabels: el.showDataLabels,
    valMin: el.valMin,
    valMax: el.valMax,
    catAxisTitle: el.catAxisTitle ?? null,
    valAxisTitle: el.valAxisTitle ?? null,
    catAxisTitleFontSizeHpt: el.catAxisTitleSize ?? null,
    catAxisTitleFontBold: el.catAxisTitleBold ?? null,
    catAxisTitleFontColor: el.catAxisTitleColor ?? null,
    valAxisTitleFontSizeHpt: el.valAxisTitleSize ?? null,
    valAxisTitleFontBold: el.valAxisTitleBold ?? null,
    valAxisTitleFontColor: el.valAxisTitleColor ?? null,
    catAxisHidden: el.catAxisHidden,
    valAxisHidden: el.valAxisHidden,
    catAxisLineHidden: el.catAxisLineHidden ?? false,
    valAxisLineHidden: el.valAxisLineHidden ?? false,
    catAxisFontColor: el.catAxisFontColor ?? null,
    valAxisFontColor: el.valAxisFontColor ?? null,
    catAxisLineColor: el.catAxisLineColor ?? null,
    catAxisLineWidthEmu: el.catAxisLineWidthEmu ?? null,
    valAxisLineColor: el.valAxisLineColor ?? null,
    valAxisLineWidthEmu: el.valAxisLineWidthEmu ?? null,
    plotAreaBg: el.plotAreaBg,
    chartBg: el.chartBg,
    showLegend: el.showLegend,
    legendPos: el.legendPos ?? null,
    catAxisCrossBetween: el.catAxisCrossBetween,
    valAxisMajorTickMark: el.valAxisMajorTickMark,
    catAxisMajorTickMark: el.catAxisMajorTickMark,
    titleFontSizeHpt: el.titleFontSizeHpt,
    titleFontColor: el.titleFontColor ?? null,
    titleFontFace: el.titleFontFace ?? null,
    titleFontBold: el.titleFontBold ?? null,
    catAxisFontSizeHpt: el.catAxisFontSizeHpt,
    valAxisFontSizeHpt: el.valAxisFontSizeHpt,
    catAxisFontBold: el.catAxisFontBold ?? null,
    valAxisFontBold: el.valAxisFontBold ?? null,
    dataLabelFontSizeHpt: el.dataLabelFontSizeHpt,
    subtotalIndices: el.subtotalIndices,
    barGapWidth: el.barGapWidth ?? null,
    barOverlap: el.barOverlap ?? null,
    dataLabelPosition: el.dataLabelPosition ?? null,
    dataLabelFontColor: el.dataLabelFontColor ?? null,
    dataLabelFormatCode: el.dataLabelFormatCode ?? null,
    valAxisFormatCode: el.valAxisFormatCode ?? null,
    plotAreaManualLayout: el.plotAreaManualLayout ?? null,
    scatterStyle: el.scatterStyle ?? null,
    radarStyle: el.radarStyle ?? null,
    chartBorderColor: el.chartBorderColor ?? null,
    chartBorderWidthEmu: el.chartBorderWidthEmu ?? null,
    secondaryValAxis: el.secondaryValAxis ?? null,
  };
}

/** Drop object properties whose value is null/undefined; recurse into objects
 *  and arrays. Array ELEMENTS are never dropped (a null data point stays). */
function stripNullish(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(stripNullish);
  if (v && typeof v === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, val] of Object.entries(v as Record<string, unknown>)) {
      if (val === null || val === undefined) continue;
      out[k] = stripNullish(val);
    }
    return out;
  }
  return v;
}

/**
 * Build the ChartModel the Rust parser now emits for the same source, applying
 * the exact field-name map the parser uses (axis-title run-prop fields renamed
 * to the core names; every other field passed through). This mirrors what
 * `el.chart` contains at render time.
 */
function nestedRustChart(el: LegacyPptxChartElement): ChartModel {
  return {
    chartType: el.chartType,
    title: el.title,
    categories: el.categories,
    series: el.series,
    valMax: el.valMax,
    valMin: el.valMin,
    subtotalIndices: el.subtotalIndices,
    showDataLabels: el.showDataLabels,
    catAxisHidden: el.catAxisHidden,
    valAxisHidden: el.valAxisHidden,
    plotAreaBg: el.plotAreaBg,
    chartBg: el.chartBg,
    showLegend: el.showLegend,
    catAxisCrossBetween: el.catAxisCrossBetween,
    valAxisMajorTickMark: el.valAxisMajorTickMark,
    catAxisMajorTickMark: el.catAxisMajorTickMark,
    titleFontSizeHpt: el.titleFontSizeHpt,
    titleFontColor: el.titleFontColor ?? null,
    titleFontFace: el.titleFontFace ?? null,
    catAxisFontSizeHpt: el.catAxisFontSizeHpt,
    valAxisFontSizeHpt: el.valAxisFontSizeHpt,
    catAxisFontColor: el.catAxisFontColor ?? null,
    valAxisFontColor: el.valAxisFontColor ?? null,
    catAxisLineColor: el.catAxisLineColor ?? null,
    catAxisLineWidthEmu: el.catAxisLineWidthEmu ?? null,
    catAxisLineHidden: el.catAxisLineHidden ?? false,
    valAxisLineColor: el.valAxisLineColor ?? null,
    valAxisLineWidthEmu: el.valAxisLineWidthEmu ?? null,
    valAxisLineHidden: el.valAxisLineHidden ?? false,
    dataLabelFontSizeHpt: el.dataLabelFontSizeHpt,
    legendPos: el.legendPos ?? null,
    barGapWidth: el.barGapWidth ?? null,
    barOverlap: el.barOverlap ?? null,
    dataLabelPosition: el.dataLabelPosition ?? null,
    dataLabelFontColor: el.dataLabelFontColor ?? null,
    dataLabelFormatCode: el.dataLabelFormatCode ?? null,
    valAxisFormatCode: el.valAxisFormatCode ?? null,
    plotAreaManualLayout: el.plotAreaManualLayout ?? null,
    scatterStyle: el.scatterStyle ?? null,
    catAxisTitle: el.catAxisTitle ?? null,
    valAxisTitle: el.valAxisTitle ?? null,
    // The six renamed fields.
    catAxisTitleFontSizeHpt: el.catAxisTitleSize ?? null,
    catAxisTitleFontBold: el.catAxisTitleBold ?? null,
    catAxisTitleFontColor: el.catAxisTitleColor ?? null,
    valAxisTitleFontSizeHpt: el.valAxisTitleSize ?? null,
    valAxisTitleFontBold: el.valAxisTitleBold ?? null,
    valAxisTitleFontColor: el.valAxisTitleColor ?? null,
    titleFontBold: el.titleFontBold ?? null,
    catAxisFontBold: el.catAxisFontBold ?? null,
    valAxisFontBold: el.valAxisFontBold ?? null,
    chartBorderColor: el.chartBorderColor ?? null,
    chartBorderWidthEmu: el.chartBorderWidthEmu ?? null,
    secondaryValAxis: el.secondaryValAxis ?? null,
  };
}

/** A representative fully-populated combo chart exercising every mapped field
 *  (renamed axis titles, secondary axis, series with markers + a null point). */
const richFixture: LegacyPptxChartElement = {
  chartType: 'clusteredBar',
  title: 'Revenue vs. margin',
  categories: ['Q1', 'Q2', 'Q3'],
  series: [
    {
      name: 'Revenue',
      color: '4472C4',
      values: [120, null, 240],
      seriesType: null,
      useSecondaryAxis: null,
    },
    {
      name: 'Margin',
      color: 'ED7D31',
      values: [0.31, 0.4, 0.28],
      seriesType: 'line',
      useSecondaryAxis: true,
      showMarker: true,
      markerSymbol: 'circle',
    },
  ],
  showDataLabels: true,
  valMin: 0,
  valMax: 300,
  catAxisTitle: 'Quarter',
  valAxisTitle: 'Revenue ($M)',
  catAxisTitleSize: 1000,
  catAxisTitleBold: true,
  catAxisTitleColor: '404040',
  valAxisTitleSize: 1100,
  valAxisTitleBold: false,
  valAxisTitleColor: '595959',
  catAxisHidden: false,
  valAxisHidden: false,
  catAxisLineHidden: true,
  valAxisLineHidden: false,
  catAxisFontColor: '808080',
  valAxisFontColor: null,
  catAxisLineColor: 'D9D9D9',
  catAxisLineWidthEmu: 9525,
  valAxisLineColor: null,
  valAxisLineWidthEmu: null,
  plotAreaBg: null,
  chartBg: 'FFFFFF',
  showLegend: true,
  legendPos: 'b',
  catAxisCrossBetween: 'between',
  valAxisMajorTickMark: 'out',
  catAxisMajorTickMark: 'out',
  titleFontSizeHpt: 1400,
  titleFontColor: '1B4332',
  titleFontFace: 'Calibri',
  titleFontBold: true,
  catAxisFontSizeHpt: 900,
  valAxisFontSizeHpt: 900,
  catAxisFontBold: false,
  valAxisFontBold: false,
  dataLabelFontSizeHpt: 800,
  subtotalIndices: [],
  barGapWidth: 150,
  barOverlap: -27,
  dataLabelPosition: 'outEnd',
  dataLabelFontColor: '000000',
  dataLabelFormatCode: '#,##0',
  valAxisFormatCode: '$#,##0',
  plotAreaManualLayout: { xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.1, w: 0.8, h: 0.7 },
  scatterStyle: null,
  radarStyle: null,
  chartBorderColor: '1B4332',
  chartBorderWidthEmu: 19050,
  secondaryValAxis: {
    min: 0,
    max: 1,
    title: 'Margin (%)',
    hidden: false,
    formatCode: '0%',
    lineHidden: false,
    majorTickMark: 'out',
  },
};

/** A sparse chart (mostly defaults / nulls) — the common single-series case. */
const sparseFixture: LegacyPptxChartElement = {
  chartType: 'pie',
  title: null,
  categories: ['A', 'B', 'C'],
  series: [{ name: '', color: null, values: [10, 20, 30] }],
  showDataLabels: false,
  valMin: null,
  valMax: null,
  catAxisHidden: false,
  valAxisHidden: false,
  plotAreaBg: null,
  chartBg: null,
  showLegend: false,
  catAxisCrossBetween: 'between',
  valAxisMajorTickMark: 'cross',
  catAxisMajorTickMark: 'cross',
  titleFontSizeHpt: null,
  catAxisFontSizeHpt: null,
  valAxisFontSizeHpt: null,
  dataLabelFontSizeHpt: null,
  subtotalIndices: [],
};

describe('C10 pptx chart nested emit — oracle deep-equal', () => {
  for (const [label, fixture] of [
    ['rich combo chart', richFixture],
    ['sparse pie chart', sparseFixture],
  ] as const) {
    it(`${label}: nested Rust emit renders identically to the frozen legacy literal`, () => {
      const expected = stripNullish(legacyPptxChartModel(fixture));
      const actual = stripNullish(nestedRustChart(fixture));
      expect(actual).toEqual(expected);
    });
  }

  it('preserves null data points inside series.values (array elements not stripped)', () => {
    const out = stripNullish(nestedRustChart(richFixture)) as ChartModel;
    expect(out.series[0].values).toEqual([120, null, 240]);
  });
});
