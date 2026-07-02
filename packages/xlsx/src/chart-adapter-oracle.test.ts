import { describe, it, expect } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';

// ─────────────────────────────────────────────────────────────────────────────
// Oracle for C10 — xlsx chart adapter moved to Rust.
//
// The renderer's `adaptChartData` + `canonicalChartType` (which normalized the
// parser-native `ChartData` into a core `ChartModel`) were deleted; the Rust
// parser now emits `ChartModel` directly via `From<ChartData>`. This test
// freezes the DELETED TS adapter (verbatim) as the oracle and proves the Rust
// conversion is equivalent: for the same `ChartData`, the object the parser now
// emits (reconstructed here by `rustFromChartData`, applying the exact rules the
// Rust `From` impl uses) deep-equals the frozen adapter output.
//
// `null` vs missing key is normalized away (`stripNullish`) — the core renderer
// reads every optional field as `?? default` / `!= null`, so absent and null
// render identically. Array elements (null data points) are preserved.
// ─────────────────────────────────────────────────────────────────────────────

/** Minimal shape of the parser-native `ChartData` (the fields the adapter reads). */
interface XlsxChartDataLike {
  chartType: string;
  barDir: string;
  grouping: string;
  title: string | null;
  categories: string[];
  series: Array<{
    name: string;
    seriesType: string;
    categories: string[];
    values: (number | null)[];
    color?: string | null;
    showMarker?: boolean;
    valFormatCode?: string | null;
    labelColor?: string | null;
    order?: number;
    markerSymbol?: string;
    markerSize?: number;
    markerFill?: string;
    markerLine?: string;
    dataPointOverrides?: unknown[];
    dataLabelOverrides?: unknown[];
    seriesDataLabels?: unknown;
    errBars?: unknown[];
  }>;
  showDataLabels?: boolean;
  catAxisTitle?: string | null;
  valAxisTitle?: string | null;
  catAxisTitleSize?: number | null;
  catAxisTitleBold?: boolean | null;
  catAxisTitleColor?: string | null;
  valAxisTitleSize?: number | null;
  valAxisTitleBold?: boolean | null;
  valAxisTitleColor?: string | null;
  showLegend?: boolean;
  legendPos?: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  titleFontSizeHpt?: number | null;
  titleFontColor?: string | null;
  titleFontFace?: string | null;
  catAxisFontSizeHpt?: number | null;
  valAxisFontSizeHpt?: number | null;
  chartBg?: string | null;
  hasChartSpPr?: boolean;
  legendManualLayout?: unknown;
  catAxisHidden?: boolean;
  valAxisHidden?: boolean;
  catAxisLineHidden?: boolean;
  valAxisLineHidden?: boolean;
  radarStyle?: string | null;
  valAxisFormatCode?: string | null;
  barGapWidth?: number | null;
  barOverlap?: number | null;
  dataLabelPosition?: string | null;
  dataLabelFontColor?: string | null;
  dataLabelFormatCode?: string | null;
  titleFontBold?: boolean;
  catAxisFontBold?: boolean;
  valAxisFontBold?: boolean;
  catAxisCrosses?: string;
  catAxisCrossesAt?: number;
  valAxisCrosses?: string;
  valAxisCrossesAt?: number;
  catAxisLineColor?: string;
  catAxisLineWidthEmu?: number;
  valAxisLineColor?: string;
  valAxisLineWidthEmu?: number;
  chartBorderColor?: string | null;
  chartBorderWidthEmu?: number | null;
  catAxisMajorTickMark?: string;
  catAxisMinorTickMark?: string;
  valAxisMajorTickMark?: string;
  valAxisMinorTickMark?: string;
  catAxisFormatCode?: string;
  catAxisMin?: number;
  catAxisMax?: number;
  valAxisMin?: number;
  valAxisMax?: number;
  titleManualLayout?: unknown;
  plotAreaManualLayout?: unknown;
}

/** FROZEN copy of the deleted `canonicalChartType` (xlsx renderer, pre-C10). */
function canonicalChartType(chart: XlsxChartDataLike): string {
  const t = chart.chartType;
  const g = chart.grouping;
  if (t === 'bar') {
    const isH = chart.barDir === 'bar';
    if (g === 'stacked')        return isH ? 'stackedBarH'    : 'stackedBar';
    if (g === 'percentStacked') return isH ? 'stackedBarHPct' : 'stackedBarPct';
    return isH ? 'clusteredBarH' : 'clusteredBar';
  }
  if (t === 'line') {
    if (g === 'stacked')        return 'stackedLine';
    if (g === 'percentStacked') return 'stackedLinePct';
    return 'line';
  }
  if (t === 'area') {
    if (g === 'stacked')        return 'stackedArea';
    if (g === 'percentStacked') return 'stackedAreaPct';
    return 'area';
  }
  return t;
}

/** FROZEN copy of the deleted `adaptChartData` (xlsx renderer, pre-C10). This
 *  is the oracle — do not modify it. */
function adaptChartData(chart: XlsxChartDataLike): ChartModel {
  return {
    chartType: canonicalChartType(chart),
    title: chart.title,
    categories: chart.categories,
    catAxisFormatCode: chart.catAxisFormatCode ?? null,
    catAxisMin: chart.catAxisMin ?? null,
    catAxisMax: chart.catAxisMax ?? null,
    titleFontBold: chart.titleFontBold ?? null,
    catAxisFontBold: chart.catAxisFontBold ?? null,
    valAxisFontBold: chart.valAxisFontBold ?? null,
    catAxisCrosses: chart.catAxisCrosses ?? null,
    catAxisCrossesAt: chart.catAxisCrossesAt ?? null,
    valAxisCrosses: chart.valAxisCrosses ?? null,
    valAxisCrossesAt: chart.valAxisCrossesAt ?? null,
    catAxisLineColor: chart.catAxisLineColor ?? null,
    catAxisLineWidthEmu: chart.catAxisLineWidthEmu ?? null,
    valAxisLineColor: chart.valAxisLineColor ?? null,
    valAxisLineWidthEmu: chart.valAxisLineWidthEmu ?? null,
    series: chart.series.map(s => ({
      name: s.name,
      color: s.color ?? null,
      values: s.values,
      seriesType: s.seriesType ?? null,
      categories: s.categories.length > 0 ? s.categories : null,
      showMarker: s.showMarker ?? null,
      valFormatCode: s.valFormatCode ?? null,
      labelColor: s.labelColor ?? null,
      markerSymbol: s.markerSymbol ?? null,
      markerSize: s.markerSize ?? null,
      markerFill: s.markerFill ?? null,
      markerLine: s.markerLine ?? null,
      dataPointOverrides: (s.dataPointOverrides as ChartModel['series'][number]['dataPointOverrides']) ?? null,
      dataLabelOverrides: (s.dataLabelOverrides as ChartModel['series'][number]['dataLabelOverrides']) ?? null,
      seriesDataLabels: (s.seriesDataLabels as ChartModel['series'][number]['seriesDataLabels']) ?? null,
      errBars: (s.errBars as ChartModel['series'][number]['errBars']) ?? null,
    })),
    showDataLabels: chart.showDataLabels ?? false,
    valMin: chart.valAxisMin ?? null,
    valMax: chart.valAxisMax ?? null,
    catAxisTitle: chart.catAxisTitle ?? null,
    valAxisTitle: chart.valAxisTitle ?? null,
    catAxisTitleFontSizeHpt: chart.catAxisTitleSize ?? null,
    catAxisTitleFontBold: chart.catAxisTitleBold ?? null,
    catAxisTitleFontColor: chart.catAxisTitleColor ?? null,
    valAxisTitleFontSizeHpt: chart.valAxisTitleSize ?? null,
    valAxisTitleFontBold: chart.valAxisTitleBold ?? null,
    valAxisTitleFontColor: chart.valAxisTitleColor ?? null,
    catAxisHidden: chart.catAxisHidden ?? false,
    valAxisHidden: chart.valAxisHidden ?? false,
    catAxisLineHidden: chart.catAxisLineHidden ?? false,
    valAxisLineHidden: chart.valAxisLineHidden ?? false,
    plotAreaBg: null,
    chartBg: chart.hasChartSpPr ? (chart.chartBg ?? null) : 'FFFFFF',
    legendManualLayout: (chart.legendManualLayout as ChartModel['legendManualLayout']) ?? null,
    showLegend: chart.showLegend ?? false,
    legendPos: chart.legendPos ?? null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: chart.valAxisMajorTickMark ?? 'out',
    catAxisMajorTickMark: chart.catAxisMajorTickMark ?? 'out',
    valAxisMinorTickMark: chart.valAxisMinorTickMark ?? null,
    catAxisMinorTickMark: chart.catAxisMinorTickMark ?? null,
    titleFontSizeHpt: chart.titleFontSizeHpt ?? null,
    titleFontColor: chart.titleFontColor ?? null,
    titleFontFace: chart.titleFontFace ?? null,
    catAxisFontSizeHpt: chart.catAxisFontSizeHpt ?? null,
    valAxisFontSizeHpt: chart.valAxisFontSizeHpt ?? null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisFormatCode: chart.valAxisFormatCode ?? null,
    barGapWidth: chart.barGapWidth ?? null,
    barOverlap: chart.barOverlap ?? null,
    dataLabelPosition: chart.dataLabelPosition ?? null,
    dataLabelFontColor: chart.dataLabelFontColor ?? null,
    dataLabelFormatCode: chart.dataLabelFormatCode ?? null,
    titleManualLayout: (chart.titleManualLayout as ChartModel['titleManualLayout']) ?? null,
    plotAreaManualLayout: (chart.plotAreaManualLayout as ChartModel['plotAreaManualLayout']) ?? null,
    radarStyle: chart.radarStyle ?? null,
    chartBorderColor: chart.chartBorderColor ?? null,
    chartBorderWidthEmu: chart.chartBorderWidthEmu ?? null,
  };
}

/**
 * Reconstruct what the Rust `From<ChartData> for ChartModel` now emits — the
 * SAME rules, implemented independently here so the two must agree. Any drift
 * between the Rust impl and the frozen adapter surfaces as a failed deep-equal.
 */
function rustFromChartData(c: XlsxChartDataLike): ChartModel {
  const chartBg = c.hasChartSpPr ? (c.chartBg ?? null) : 'FFFFFF';
  return {
    chartType: canonicalChartType(c), // canonical_chart_type() in Rust — same table
    title: c.title,
    categories: c.categories,
    series: c.series.map(s => ({
      name: s.name,
      color: s.color ?? null,
      values: s.values,
      dataPointColors: null,
      dataLabelColors: null,
      labelColor: s.labelColor ?? null,
      seriesType: s.seriesType,
      useSecondaryAxis: null,
      categories: s.categories.length > 0 ? s.categories : null,
      showMarker: s.showMarker ?? null,
      valFormatCode: s.valFormatCode ?? null,
      markerSymbol: s.markerSymbol ?? null,
      markerSize: s.markerSize ?? null,
      markerFill: s.markerFill ?? null,
      markerLine: s.markerLine ?? null,
      dataPointOverrides: (s.dataPointOverrides && s.dataPointOverrides.length > 0
        ? s.dataPointOverrides : null) as ChartModel['series'][number]['dataPointOverrides'],
      dataLabelOverrides: (s.dataLabelOverrides && s.dataLabelOverrides.length > 0
        ? s.dataLabelOverrides : null) as ChartModel['series'][number]['dataLabelOverrides'],
      seriesDataLabels: (s.seriesDataLabels as ChartModel['series'][number]['seriesDataLabels']) ?? null,
      errBars: (s.errBars && s.errBars.length > 0
        ? s.errBars : null) as ChartModel['series'][number]['errBars'],
      bubbleSizes: null,
    })),
    showDataLabels: c.showDataLabels ?? false,
    valMin: c.valAxisMin ?? null,
    valMax: c.valAxisMax ?? null,
    catAxisTitle: c.catAxisTitle ?? null,
    valAxisTitle: c.valAxisTitle ?? null,
    catAxisHidden: c.catAxisHidden ?? false,
    valAxisHidden: c.valAxisHidden ?? false,
    catAxisLineHidden: c.catAxisLineHidden ?? false,
    valAxisLineHidden: c.valAxisLineHidden ?? false,
    plotAreaBg: null,
    chartBg,
    showLegend: c.showLegend ?? false,
    legendPos: c.legendPos ?? null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: c.valAxisMajorTickMark ?? 'out',
    catAxisMajorTickMark: c.catAxisMajorTickMark ?? 'out',
    titleFontSizeHpt: c.titleFontSizeHpt ?? null,
    titleFontColor: c.titleFontColor ?? null,
    titleFontFace: c.titleFontFace ?? null,
    catAxisFontSizeHpt: c.catAxisFontSizeHpt ?? null,
    valAxisFontSizeHpt: c.valAxisFontSizeHpt ?? null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisMinorTickMark: c.valAxisMinorTickMark ?? null,
    catAxisMinorTickMark: c.catAxisMinorTickMark ?? null,
    catAxisFontColor: null,
    valAxisFontColor: null,
    legendManualLayout: (c.legendManualLayout as ChartModel['legendManualLayout']) ?? null,
    valAxisFormatCode: c.valAxisFormatCode ?? null,
    barGapWidth: c.barGapWidth ?? null,
    barOverlap: c.barOverlap ?? null,
    dataLabelPosition: c.dataLabelPosition ?? null,
    dataLabelFontColor: c.dataLabelFontColor ?? null,
    dataLabelFormatCode: c.dataLabelFormatCode ?? null,
    titleFontBold: c.titleFontBold ?? null,
    catAxisFontBold: c.catAxisFontBold ?? null,
    valAxisFontBold: c.valAxisFontBold ?? null,
    catAxisTitleFontSizeHpt: c.catAxisTitleSize ?? null,
    catAxisTitleFontBold: c.catAxisTitleBold ?? null,
    catAxisTitleFontColor: c.catAxisTitleColor ?? null,
    valAxisTitleFontSizeHpt: c.valAxisTitleSize ?? null,
    valAxisTitleFontBold: c.valAxisTitleBold ?? null,
    valAxisTitleFontColor: c.valAxisTitleColor ?? null,
    chartBorderColor: c.chartBorderColor ?? null,
    chartBorderWidthEmu: c.chartBorderWidthEmu ?? null,
    catAxisCrosses: c.catAxisCrosses ?? null,
    catAxisCrossesAt: c.catAxisCrossesAt ?? null,
    valAxisCrosses: c.valAxisCrosses ?? null,
    valAxisCrossesAt: c.valAxisCrossesAt ?? null,
    catAxisLineColor: c.catAxisLineColor ?? null,
    catAxisLineWidthEmu: c.catAxisLineWidthEmu ?? null,
    valAxisLineColor: c.valAxisLineColor ?? null,
    valAxisLineWidthEmu: c.valAxisLineWidthEmu ?? null,
    catAxisFormatCode: c.catAxisFormatCode ?? null,
    catAxisMin: c.catAxisMin ?? null,
    catAxisMax: c.catAxisMax ?? null,
    titleManualLayout: (c.titleManualLayout as ChartModel['titleManualLayout']) ?? null,
    plotAreaManualLayout: (c.plotAreaManualLayout as ChartModel['plotAreaManualLayout']) ?? null,
    scatterStyle: null,
    radarStyle: c.radarStyle ?? null,
    secondaryValAxis: null,
  };
}

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

const richBar: XlsxChartDataLike = {
  chartType: 'bar',
  barDir: 'bar',
  grouping: 'stacked',
  title: 'Regional split',
  categories: ['North', 'South'],
  series: [
    {
      name: 'Units',
      seriesType: 'bar',
      categories: ['North', 'South'],
      values: [10, null],
      color: '4472C4',
      showMarker: false,
      order: 0,
      valFormatCode: '#,##0',
      dataPointOverrides: [{ idx: 0, color: 'FF0000' }],
    },
  ],
  showDataLabels: true,
  catAxisTitle: 'Region',
  valAxisTitle: 'Units',
  catAxisTitleSize: 1000,
  catAxisTitleBold: true,
  catAxisTitleColor: '404040',
  showLegend: true,
  legendPos: 'b',
  titleFontSizeHpt: 1400,
  catAxisFontSizeHpt: 900,
  valAxisFontSizeHpt: 900,
  hasChartSpPr: true,
  chartBg: 'F2F2F2',
  catAxisHidden: false,
  valAxisHidden: false,
  radarStyle: null,
  valAxisFormatCode: '#,##0',
  barGapWidth: 150,
  barOverlap: 100,
  catAxisMajorTickMark: 'out',
  valAxisMajorTickMark: 'none',
  valAxisMin: 0,
  valAxisMax: 40,
  catAxisCrosses: 'autoZero',
};

/** Single-series vertical bar, no spPr → white-default chartBg; no categories.
 *  `showMarker` is `false` (never omitted): the Rust `ChartData.show_marker` is
 *  a non-optional `bool`, so the real parser JSON always carries it — and the
 *  old adapter passed it through (`false ?? null` = `false`). Omitting it here
 *  would model input the parser never actually produces. */
const defaultsBar: XlsxChartDataLike = {
  chartType: 'bar',
  barDir: 'col',
  grouping: 'clustered',
  title: null,
  categories: [],
  series: [{ name: 'S', seriesType: 'bar', categories: [], values: [1, 2, 3], showMarker: false, order: 0 }],
};

describe('C10 xlsx chart adapter → Rust — oracle deep-equal', () => {
  for (const [label, data] of [
    ['rich horizontal stacked bar', richBar],
    ['defaults-only vertical bar (white bg, no cats)', defaultsBar],
  ] as const) {
    it(`${label}: Rust From<ChartData> matches the frozen adapter`, () => {
      const expected = stripNullish(adaptChartData(data));
      const actual = stripNullish(rustFromChartData(data));
      expect(actual).toEqual(expected);
    });
  }

  it('white-default vs noFill chartBg branch', () => {
    // No spPr → default opaque white.
    expect((rustFromChartData(defaultsBar) as ChartModel).chartBg).toBe('FFFFFF');
    // spPr present, chartBg resolved null (noFill) → transparent, not white.
    const noFill: XlsxChartDataLike = { ...defaultsBar, hasChartSpPr: true, chartBg: null };
    expect((rustFromChartData(noFill) as ChartModel).chartBg).toBeNull();
  });

  it('canonicalizes bar direction + grouping like the Rust table', () => {
    const mk = (barDir: string, grouping: string) =>
      canonicalChartType({ ...defaultsBar, barDir, grouping });
    expect(mk('col', 'clustered')).toBe('clusteredBar');
    expect(mk('bar', 'clustered')).toBe('clusteredBarH');
    expect(mk('bar', 'percentStacked')).toBe('stackedBarHPct');
  });
});
