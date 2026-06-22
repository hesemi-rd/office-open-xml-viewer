// Unified chart renderer. Dispatches on canonical `ChartModel.chartType` and
// delegates to per-family implementations (bar, line, area, pie, radar,
// scatter, waterfall). Ported from the xlsx implementation with pptx
// extensions (valMin-aware axis, plotAreaBg, dataPointColors, waterfall).

import type { ChartModel, ChartRect, ChartSeries } from '../types/chart';
import { niceStep, valueAxisScale } from './axis-scale.js';
import { formatChartVal, formatChartValWithCode } from './chart-number-format.js';
import { hexToRgba } from '../shape/paint.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';

// ─── Palette + helpers ──────────────────────────────────────────────────────

export const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];

function chartColor(idx: number, series?: { color?: string | null } | null): string {
  if (series?.color) return `#${series.color}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

function pieSliceColor(idx: number, series: ChartSeries): string {
  const override = series.dataPointColors?.[idx];
  if (override) return `#${override}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

/** Chart types whose legend lists one entry per category (data point of the
 *  first series) rather than one entry per series. Excel/PowerPoint draw pie
 *  and doughnut legends this way: each slice gets its own row, colored with
 *  the slice's color. ECMA-376 §21.2.2.114 (`<c:varyColors>` defaults true for
 *  pie/doughnut). */
function legendIsCategoryDriven(chartType: string | undefined): boolean {
  return chartType === 'pie' || chartType === 'doughnut';
}

/** Resolve the color for legend entry `entryIndex`, matching the marks the
 *  plot actually draws.
 *
 *  - Category-driven legends (pie / doughnut): the entry maps to data point
 *    `entryIndex` of the first series, so it must use the *same* resolution as
 *    {@link pieSliceColor} — explicit per-point `dPt` color, else the palette
 *    indexed by point. The series-level fill is deliberately ignored: a pie
 *    series carries a single `<c:spPr>` solidFill that, if honored here, would
 *    collapse every swatch to one color while the slices stay multi-colored.
 *  - Series-driven legends (bar / line / area / …): the entry maps to series
 *    `entryIndex`, so it uses {@link chartColor} — explicit series fill else
 *    the palette indexed by series. */
export function legendEntryColor(
  chartType: string | undefined,
  series: ChartSeries[],
  entryIndex: number,
): string {
  if (legendIsCategoryDriven(chartType)) {
    const first = series[0];
    if (first) return pieSliceColor(entryIndex, first);
    return `#${CHART_PALETTE[entryIndex % CHART_PALETTE.length]}`;
  }
  return chartColor(entryIndex, series[entryIndex]);
}

/** Axis-title font size (px). Honors the XML run-prop size (`<c:catAx|valAx>
 *  <c:title>…@sz`, hundredths of a point) when present — e.g. 18pt titles in
 *  sample-30 — otherwise keeps the previous proportional default so untitled-
 *  size charts are unchanged. */
function axisTitleFontPx(sizeHpt: number | null | undefined, h: number, ptToPx: number): number {
  if (sizeHpt) return (sizeHpt / 100) * ptToPx;
  return Math.max(8, Math.min(10, h * 0.045));
}

/** Draw an axis title at an explicit anchor in the outer gutter band (outside
 *  the tick labels), at its real font size/bold/color. The cat title is
 *  centered under the X axis; the val title is rotated -90° centered to the
 *  left of the Y axis. `anchorX`/`anchorY` are the band center the caller
 *  reserved via catTitleH/valTitleW, so the title never overlaps tick labels. */
function drawAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchorX: number, anchorY: number,
  axis: 'cat' | 'val',
  fontSizePx: number,
  bold: boolean,
  color: string,
): void {
  ctx.save();
  ctx.font = `${bold ? 'bold ' : ''}${fontSizePx}px sans-serif`;
  ctx.fillStyle = color;
  if (axis === 'cat') {
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(text.slice(0, 60), anchorX, anchorY);
  } else {
    ctx.translate(anchorX, anchorY);
    ctx.rotate(-Math.PI / 2);
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(text.slice(0, 60), 0, 0);
  }
  ctx.restore();
}

/** Resolve the per-axis title color string for `drawAxisTitle`. Returns
 *  '#rrggbb' when the XML supplied a srgb color, else the legacy '#555'. */
function axisTitleColor(hex: string | null | undefined): string {
  return hex ? `#${hex}` : '#555';
}

/** Margin (px) between the chart's outer edge and an axis title. ECMA-376
 *  leaves axis-title position automatic when there is no `<c:layout>/<c:manualLayout>`
 *  (§21.2.2.88 layout / §21.2.2.32 plotArea) — there is NO spec-defined inset, so
 *  the `8px` floor and `2%` factor are empirical approximations of Excel's
 *  auto-layout, not measured spec values. `dim` = width (left/val title) or
 *  height (bottom/cat title). */
function axisTitleMargin(dim: number): number {
  return Math.max(8, dim * 0.02);
}

interface AxisTitleLayout { catTitlePx: number; valTitlePx: number; catTitleH: number; valTitleW: number; }

/** Axis-title font sizes (px) and the outer gutter bands to reserve for them
 *  (`catTitleH` → bottom pad, `valTitleW` → left pad). Identical across all four
 *  cartesian renderers; keeping reserve here and the draw anchors in
 *  drawAxisTitles in sync via one source for the band size. The per-renderer
 *  `pad` formula that consumes these differs and stays inline. */
function axisTitleLayout(chart: ChartModel, w: number, h: number, ptToPx: number): AxisTitleLayout {
  const catTitlePx = axisTitleFontPx(chart.catAxisTitleFontSizeHpt, h, ptToPx);
  const valTitlePx = axisTitleFontPx(chart.valAxisTitleFontSizeHpt, h, ptToPx);
  return {
    catTitlePx, valTitlePx,
    catTitleH: chart.catAxisTitle ? catTitlePx + axisTitleMargin(h) + 4 : 0,
    valTitleW: chart.valAxisTitle ? valTitlePx + axisTitleMargin(w) + 4 : 0,
  };
}

/** Draw both axis titles for a cartesian chart (bar/line/area/scatter),
 *  anchored in the reserved outer gutter bands so they sit OUTSIDE the tick
 *  labels. `catTitlePx`/`valTitlePx` are the title font sizes the caller used
 *  to size `catTitleH`/`valTitleW`; the anchor centers each title within its
 *  band. cat axis = bottom, val axis = left — the orientation each cartesian
 *  renderer already uses (horizontal bar keeps cat-bottom/val-left too).
 *  Axis titles default to BOLD — ECMA-376 Part 1 (ST_Style, chart-style
 *  defaults) states "Axis titles and chart titles are bold by default, while
 *  all other chart elements are normal" (same clause sets the default size to
 *  10pt). So an unspecified weight renders bold; only an explicit `b="0"`
 *  un-bolds. Consistent with drawChartTitle, which applies the same default. */
function drawAxisTitles(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  legLeftW: number, legBottomH: number,
  catTitlePx: number, valTitlePx: number,
): void {
  if (chart.valAxisTitle) {
    const anchorX = x + legLeftW + axisTitleMargin(w) + valTitlePx / 2;
    const anchorY = py0 + ph / 2;
    drawAxisTitle(
      ctx, chart.valAxisTitle, anchorX, anchorY, 'val',
      valTitlePx, chart.valAxisTitleFontBold ?? true, axisTitleColor(chart.valAxisTitleFontColor),
    );
  }
  if (chart.catAxisTitle) {
    const anchorX = px0 + pw / 2;
    const anchorY = y + h - legBottomH - axisTitleMargin(h) - catTitlePx / 2;
    drawAxisTitle(
      ctx, chart.catAxisTitle, anchorX, anchorY, 'cat',
      catTitlePx, chart.catAxisTitleFontBold ?? true, axisTitleColor(chart.catAxisTitleFontColor),
    );
  }
}

/** Line-shaped legend swatch styles match Excel's actual chart-type
 *  conventions: bar/column/area/pie use a filled rectangle ("swatch");
 *  line/radar/scatter use a horizontal line segment (the same stroke
 *  weight the series uses). Without this, line-chart legends rendered as
 *  filled squares, which read as a different chart-type marker.
 */
type LegendSwatchStyle = 'fill' | 'line';

function legendSwatchStyle(chartType: string | undefined): LegendSwatchStyle {
  if (!chartType) return 'fill';
  if (
    chartType === 'line' || chartType === 'stackedLine' || chartType === 'stackedLinePct' ||
    chartType === 'radar' || chartType === 'scatter'
  ) {
    return 'line';
  }
  return 'fill';
}

function drawLegendSwatch(
  ctx: CanvasRenderingContext2D,
  style: LegendSwatchStyle,
  color: string,
  x: number, y: number, w: number, h: number,
): void {
  ctx.fillStyle = color;
  if (style === 'line') {
    // Horizontal stroke centered vertically inside the swatch slot. 2 px
    // weight matches Excel's default 2.25 pt line at typical legend sizes.
    ctx.strokeStyle = color;
    const prevW = ctx.lineWidth;
    ctx.lineWidth = Math.max(1.5, h * 0.15);
    ctx.beginPath();
    const ly = y + h / 2;
    ctx.moveTo(x, ly);
    ctx.lineTo(x + w, ly);
    ctx.stroke();
    ctx.lineWidth = prevW;
  } else {
    ctx.fillRect(x, y, w, h);
  }
}

/** A single legend row: a label and the color of its swatch. Built so that the
 *  swatch color is resolved exactly like the mark it represents (slice / bar /
 *  line). See {@link legendEntryColor}. */
interface LegendEntry { label: string; color: string }

/** Build the legend entries for a chart. Pie/doughnut legends are
 *  category-driven (one row per data point of the first series, labeled by
 *  category); every other chart type is series-driven (one row per series). */
function buildLegendEntries(series: ChartSeries[], chartType: string | undefined): LegendEntry[] {
  if (legendIsCategoryDriven(chartType)) {
    const first = series[0];
    const n = first ? first.values.length : 0;
    const cats = first?.categories ?? [];
    return Array.from({ length: n }, (_, i) => ({
      label: (cats[i] ?? `Item ${i + 1}`).toString(),
      color: legendEntryColor(chartType, series, i),
    }));
  }
  return series.map((s, i) => ({
    label: s.name || `Series ${i + 1}`,
    color: legendEntryColor(chartType, series, i),
  }));
}

function drawLegend(
  ctx: CanvasRenderingContext2D,
  series: ChartSeries[],
  lx: number, ly: number, lw: number, lh: number,
  orient: 'vertical' | 'horizontal' = 'vertical',
  chartType?: string,
): void {
  const sw = 10; const gap = 4;
  const swatchStyle = legendSwatchStyle(chartType);
  const entries = buildLegendEntries(series, chartType);
  if (orient === 'horizontal') {
    // Excel lays a bottom/top legend as a single horizontal row, centered.
    const fontSize = Math.max(9, Math.min(12, lh * 0.7));
    ctx.font = `${fontSize}px sans-serif`;
    ctx.textBaseline = 'middle';
    const itemGap = 12;
    const itemWidths = entries.map((e) => sw + gap + ctx.measureText(e.label.slice(0, 30)).width);
    const total = itemWidths.reduce((a, b) => a + b, 0) + itemGap * Math.max(0, entries.length - 1);
    let rx = lx + (lw - total) / 2;
    const ry = ly + lh / 2;
    for (let i = 0; i < entries.length; i++) {
      drawLegendSwatch(ctx, swatchStyle, entries[i].color, rx, ry - fontSize / 2, sw, fontSize);
      ctx.fillStyle = '#333'; ctx.textAlign = 'left';
      ctx.fillText(entries[i].label.slice(0, 30), rx + sw + gap, ry);
      rx += itemWidths[i] + itemGap;
    }
    return;
  }
  const fontSize = Math.max(9, Math.min(12, lh / (entries.length + 1)));
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
  const rowH = fontSize + 4;
  let ry = ly + (lh - rowH * entries.length) / 2;
  for (let i = 0; i < entries.length; i++) {
    drawLegendSwatch(ctx, swatchStyle, entries[i].color, lx, ry, sw, fontSize);
    ctx.fillStyle = '#333'; ctx.textAlign = 'left';
    ctx.fillText(entries[i].label.slice(0, 20), lx + sw + gap, ry + fontSize / 2);
    ry += rowH;
  }
  void lw;
}

type LegendSide = 'r' | 'l' | 't' | 'b';
interface LegendLayout {
  side: LegendSide;
  /** Reserved plot-area width (>0 when side = l or r). */
  reserveW: number;
  /** Reserved plot-area height (>0 when side = t or b). */
  reserveH: number;
}

/** Resolve legend placement from `<c:legendPos>`. Returns null when hidden. */
function legendLayout(chart: ChartModel, w: number, h: number): LegendLayout | null {
  if (!chart.showLegend) return null;
  const pos = chart.legendPos ?? 'r';
  const side: LegendSide = pos === 'l' ? 'l' : pos === 't' ? 't' : pos === 'b' ? 'b' : 'r';
  if (side === 'r' || side === 'l') {
    return { side, reserveW: Math.max(80, w * 0.22), reserveH: 0 };
  }
  // Excel's top/bottom legend is a single-row strip; reserve ~8% of height.
  return { side, reserveW: 0, reserveH: Math.max(18, h * 0.08) };
}

/** Draw a legend in the band reserved by {@link legendLayout}. */
function drawLegendForLayout(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  leg: LegendLayout | null,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  topBand: number,
): void {
  if (!leg) return;
  // `<c:legend><c:manualLayout>` (§21.2.2.31) wins over the default side-based
  // rectangle. We honor the `edge` placement mode — fractions are measured
  // from the top-left of the chart space — which matches what Excel's built-in
  // templates emit. `factor` mode (offset from default) is rarer; fall back to
  // the reserved band in that case rather than guess.
  const ml = chart.legendManualLayout;
  if (ml && ml.xMode === 'edge' && ml.yMode === 'edge' && ml.w > 0 && ml.h > 0) {
    const lx = x + ml.x * w;
    const ly = y + ml.y * h;
    const lw = ml.w * w;
    const lh = ml.h * h;
    // Legend is always a horizontal strip when placed on top/bottom; vertical
    // when on left/right. A manual box wider than tall implies horizontal —
    // matches Excel's one-row legend rendering for top/bottom manual layouts.
    const orient = lw >= lh ? 'horizontal' : 'vertical';
    drawLegend(ctx, chart.series, lx, ly, lw, lh, orient, chart.chartType);
    return;
  }
  switch (leg.side) {
    case 'r':
      drawLegend(ctx, chart.series, x + w - leg.reserveW + 4, py0, leg.reserveW - 8, ph, 'vertical', chart.chartType);
      break;
    case 'l':
      drawLegend(ctx, chart.series, x + 4, py0, leg.reserveW - 8, ph, 'vertical', chart.chartType);
      break;
    case 't':
      drawLegend(ctx, chart.series, px0, y + topBand, pw, leg.reserveH, 'horizontal', chart.chartType);
      break;
    case 'b':
      drawLegend(ctx, chart.series, px0, y + h - leg.reserveH, pw, leg.reserveH, 'horizontal', chart.chartType);
      break;
  }
}

function drawAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string | null | undefined,
  axis: 'val' | 'cat',
  anchorXOrY: number,
  perpendicular: number,
  color?: string,
  lineWidth?: number,
): void {
  if (mode === 'none' || !mode) return;
  // Tick length scales mildly with the axis line weight so a thick
  // ruler-style axis (e.g. Vertex42 Gantt 5 pt) produces ticks that
  // are visible without being huge.
  const baseLen = 4;
  const len = lineWidth ? Math.max(baseLen, lineWidth + 2) : baseLen;
  const prevS = ctx.strokeStyle;
  const prevW = ctx.lineWidth;
  ctx.strokeStyle = color ?? '#888';
  ctx.lineWidth = lineWidth ?? 1;
  ctx.beginPath();
  if (axis === 'val') {
    // val axis is vertical (x = anchor, y varies). Ticks extend horizontally.
    const x0 = anchorXOrY;
    const y = perpendicular;
    const outer = mode === 'out' || mode === 'cross' ? -len : 0;
    const inner = mode === 'in' || mode === 'cross' ? len : 0;
    ctx.moveTo(x0 + outer, y);
    ctx.lineTo(x0 + inner, y);
  } else {
    // cat axis is horizontal (y = anchor, x varies). Ticks extend vertically.
    const y0 = anchorXOrY;
    const xc = perpendicular;
    const outer = mode === 'out' || mode === 'cross' ? len : 0;
    const inner = mode === 'in' || mode === 'cross' ? -len : 0;
    ctx.moveTo(xc, y0 + outer);
    ctx.lineTo(xc, y0 + inner);
  }
  ctx.stroke();
  ctx.strokeStyle = prevS;
  ctx.lineWidth = prevW;
}

function chartTitleFontPx(chart: ChartModel, h: number, ptToPx: number): number {
  // Honor the XML-specified title font size (hundredths of a point) when
  // present. ptToPx is the pixels-per-point at the current slide scale, so
  // a 16pt title renders at the same proportional size as PowerPoint.
  if (chart.titleFontSizeHpt) return (chart.titleFontSizeHpt / 100) * ptToPx;
  return Math.max(10, h * 0.085);
}

/** Resolve an axis label font size (px) from <c:txPr> hpt or a proportional
 *  fallback. ptToPx comes from the host renderer (EMU/px scale at display). */
function axisLabelPx(sizeHpt: number | null | undefined, h: number, ptToPx: number): number {
  if (sizeHpt) return (sizeHpt / 100) * ptToPx;
  return Math.max(8, h * 0.045);
}

function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!chart.title) return;
  const face = chart.titleFontFace ? `"${chart.titleFontFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
  ctx.font = `${(chart.titleFontBold ?? true) ? 'bold ' : ''}${fontSize}px ${face}`;
  ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(chart.title, x + w / 2, y);
}

// ─── Category helper ────────────────────────────────────────────────────────

function chartCategories(chart: ChartModel): string[] {
  if (chart.categories.length > 0) return chart.categories;
  const first = chart.series[0];
  if (first?.categories && first.categories.length > 0) return first.categories;
  // ECMA-376 §21.2.2.24 — when <c:cat> is absent the category axis uses
  // integer values starting at 1. Fall back to the longest series so the
  // chart still renders instead of bailing out at n === 0.
  let n = 0;
  for (const s of chart.series) if (s.values.length > n) n = s.values.length;
  return n > 0 ? Array.from({ length: n }, (_, i) => String(i + 1)) : [];
}

/**
 * Draw a bar data label with the ECMA-376 §21.2.2.16 `dLblPos` semantics.
 *
 * For a vertical bar the coordinates describe the rectangle top-left + width +
 * height; for a horizontal bar they describe the bar's left-edge `bx`, top `by`,
 * length `barL`, and thickness `barW`. When `position` is "inBase" / "inEnd" /
 * "ctr" the label sits inside the bar; "outEnd" (default for clustered bars)
 * nudges the text just past the far edge. An explicit `color` overrides the
 * default dark label fill — Excel's workbook typically pairs "inBase" with a
 * white text color so labels stay readable against the bar fill.
 */
function drawBarDataLabel(
  ctx: CanvasRenderingContext2D,
  text: string,
  bx: number, by: number, barL: number, barW: number,
  orient: 'vertical' | 'horizontal',
  position: string | null,
  color: string | null,
): void {
  const pos = (position ?? 'outEnd');
  const fill = color ? `#${color}` : '#333';
  ctx.fillStyle = fill;
  if (orient === 'vertical') {
    // bx/by = top-left of bar rect (bar grows upward from by+barL toward by).
    // barL here is bar height (pixels) and barW is bar width.
    const cx = bx + barW / 2;
    if (pos === 'inBase') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      ctx.fillText(text, cx, by + barL - 2);
    } else if (pos === 'inEnd') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ctx.fillText(text, cx, by + 2);
    } else if (pos === 'ctr') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, cx, by + barL / 2);
    } else {
      // outEnd / default: just above the bar's top edge (by).
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      ctx.fillText(text, cx, by - 1);
    }
  } else {
    // Horizontal: bar grows to the right from bx.
    const cy = by + barW / 2;
    if (pos === 'inBase') {
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + 4, cy);
    } else if (pos === 'inEnd') {
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL - 4, cy);
    } else if (pos === 'ctr') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL / 2, cy);
    } else {
      // outEnd / default: just past the bar's right edge.
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL + 2, cy);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart — vertical columns + horizontal bars, clustered + stacked +
// percentStacked. Also handles mixed bar+line series (seriesType per series).
// ═══════════════════════════════════════════════════════════════════════════

function renderBarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const isH = chart.chartType === 'clusteredBarH' || chart.chartType === 'stackedBarH' || chart.chartType === 'stackedBarHPct';
  const stacked = chart.chartType.startsWith('stacked');
  const pct = chart.chartType === 'stackedBarPct' || chart.chartType === 'stackedBarHPct';

  const barSeries  = chart.series.filter(s => s.seriesType !== 'line');
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');

  // Combo charts (bar + line) may bind the line series to a SECONDARY value
  // axis drawn on the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with
  // axPos="r" / `<c:crosses val="max">`). `sec` is non-null only when both the
  // axis is declared AND at least one line series opts into it; horizontal bar
  // charts never carry one.
  const sec = !isH && chart.secondaryValAxis && lineSeries.some(s => s.useSecondaryAxis === true)
    ? chart.secondaryValAxis
    : null;

  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;

  // Honor the XML-specified title font size when present; otherwise fall back
  // to the proportional heuristic. Reserve the title band based on the actual
  // drawn height so the plot shrinks to avoid overlap.
  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.02 : 0;
  const titleBottomPad = chart.title ? h * 0.025 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  // Axis-title bands sized from the *actual* title font (honoring XML @sz, e.g.
  // sample-30's 18pt) plus a small gap, so big titles get a wide enough gutter
  // and never collide with the tick labels.
  const { catTitlePx, valTitlePx, catTitleH, valTitleW } = axisTitleLayout(chart, w, h, ptToPx);
  // Value-axis scales are computed up-front (before `pad`) so the side gutters
  // can be sized to the actual tick-label widths instead of a fixed fraction of
  // the chart width — short numeric labels otherwise leave a big empty gap
  // between the axis title and the labels (PowerPoint sizes the gutter to fit
  // the labels). The scales depend only on the series data, not on `pad`.
  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    let stackSum = 0;
    for (const s of barSeries) {
      const v = s.values[ci] ?? 0;
      if (stacked) stackSum += Math.abs(v);
      else dataMax = Math.max(dataMax, Math.abs(v));
    }
    if (stacked) dataMax = Math.max(dataMax, stackSum);
  }
  if (pct) dataMax = 100;
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;
  // Bar/column anchors the value axis at 0; ignore the returned min.
  const { max: axMax, step } = valueAxisScale(0, dataMax, undefined, chart.valMax);

  // Secondary value-axis scale (combo charts). INDEPENDENT of the primary: its
  // own "nice" major unit / gridline count (the auto unit is Excel-proprietary,
  // so it can differ from PowerPoint by one step). Explicit `<c:scaling>` wins.
  let sMin = 0, sMax = 1, sStep = 1;
  if (sec) {
    const lineVals: number[] = [];
    for (const s of lineSeries) {
      if (s.useSecondaryAxis !== true) continue;
      for (const v of s.values) if (v != null) lineVals.push(v);
    }
    const dMin = lineVals.length ? Math.min(...lineVals) : 0;
    const dMax = lineVals.length ? Math.max(...lineVals) : 1;
    const scl = valueAxisScale(Math.min(0, dMin), dMax, sec.min, sec.max);
    sMin = scl.min; sMax = scl.max; sStep = scl.step;
  }

  // Vertical pads first (independent of the side gutters) so the plot height —
  // and thus the tick-label font — is known before measuring the labels.
  const padT = titleH + legTopH + h * 0.02;
  const padB = isH
    ? (chart.valAxisHidden ? h * 0.02 : h * 0.08) + catTitleH + legBottomH
    : h * 0.14 + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const tickFontPx = Math.max(8, Math.min(11, phEst / 20));
  const prevFont = ctx.font;
  // Primary value-axis label band (column charts only; horizontal bars keep a
  // wider left band for the category labels).
  let valLabelBandW = 0;
  if (!isH && !chart.valAxisHidden) {
    ctx.font = `${tickFontPx}px sans-serif`;
    let wmax = 0;
    const vSteps = Math.round(axMax / step);
    for (let si = 0; si <= vSteps; si++) {
      const label = pct
        ? `${Math.round(si * step)}%`
        : formatChartValWithCode(si * step, chart.valAxisFormatCode);
      wmax = Math.max(wmax, ctx.measureText(label).width);
    }
    valLabelBandW = wmax + 16; // ~12px tick+gap to the axis + ~4px to the title
  }
  // Secondary value-axis label band (right edge), measured the same way.
  let secLabelBandW = 0;
  if (sec && !sec.hidden) {
    ctx.font = `${secTickFontPx}px sans-serif`;
    let wmax = 0;
    const sSteps = Math.round((sMax - sMin) / sStep);
    for (let si = 0; si <= sSteps; si++) {
      wmax = Math.max(wmax, ctx.measureText(formatChartValWithCode(sMin + si * sStep, undefined)).width);
    }
    secLabelBandW = wmax + 18;
  }
  ctx.font = prevFont;
  const secTitleBandW = sec && sec.title
    ? (sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05)) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.03 + secLabelBandW + secTitleBandW,
    b: padB,
    // Column charts: title band + measured label band, tight to the axis.
    // Horizontal bars: keep the wider left band for the category labels
    // (`c:catAx/c:delete val="1"` → no category labels, so tighten).
    l: isH
      ? (chart.catAxisHidden ? w * 0.03 : w * 0.22) + valTitleW + legLeftW
      : legLeftW + valTitleW + valLabelBandW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  // Plot-area placement: honor `<c:plotArea><c:layout><c:manualLayout>` when
  // present (ECMA-376 §21.2.2.32). Templates use this to keep bars from
  // overflowing into side annotations — sample-2 slide-16's horizontal bar
  // chart has the chart frame extending into the right-hand text column,
  // and the explicit `x=0.184, w=0.797` keeps the actual bars on the left.
  // `layoutTarget="inner"` (default) means the rectangle covers the inner
  // data region; "outer" includes axes/labels. We treat both identically
  // because the inner padding stays the same either way.
  const pml = chart.plotAreaManualLayout;
  let px0: number, py0: number, pw: number, ph: number;
  if (pml && pml.w != null && pml.h != null) {
    px0 = x + pml.x * w;
    py0 = y + pml.y * h;
    pw  = pml.w * w;
    ph  = pml.h * h;
  } else {
    px0 = x + pad.l; py0 = y + pad.t;
    pw  = w - pad.l - pad.r; ph = h - pad.t - pad.b;
  }
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // `axMax`/`step` (primary) and `sMin`/`sMax`/`sStep` (secondary) were computed
  // above the `pad` block so the gutters could be sized to the labels. The
  // line-mapping helpers need the now-final plot rect, so they live here. Line
  // series bound to the secondary axis map through `toYSecondary`; everything
  // else uses the primary `axMax`.
  const sRange = (sMax - sMin) || 1;
  const toYPrimaryLine = (v: number): number => py0 + ph - (v / axMax) * ph;
  const toYSecondary   = (v: number): number => py0 + ph - ((v - sMin) / sRange) * ph;

  const gridColor = '#e0e0e0';
  const steps = Math.round(axMax / step);
  ctx.textBaseline = 'middle';
  ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
  // Honor `<c:valAx><c:txPr>…<a:solidFill>` when present (ECMA-376 §21.2.2.*);
  // otherwise keep the neutral gray default.
  const valLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
  ctx.fillStyle = valLabelColor;

  if (!chart.valAxisHidden) {
    for (let si = 0; si <= steps; si++) {
      const val = si * step;
      const label = pct
        ? `${Math.round(val)}%`
        : formatChartValWithCode(val, chart.valAxisFormatCode);
      if (!isH) {
        const gy = py0 + ph - (val / axMax) * ph;
        ctx.strokeStyle = si === 0 ? '#aaa' : gridColor;
        ctx.lineWidth = si === 0 ? 1 : 0.5;
        ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
        ctx.textAlign = 'right';
        ctx.fillText(label, px0 - 12, gy);
      } else {
        const gx = px0 + (val / axMax) * pw;
        ctx.strokeStyle = si === 0 ? '#aaa' : gridColor;
        ctx.lineWidth = si === 0 ? 1 : 0.5;
        ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
        ctx.textAlign = 'center';
        ctx.fillText(label, gx, py0 + ph + 10);
      }
    }
  }

  // Axis rules. The CATEGORY axis runs along the bars' baseline — bottom
  // (horizontal) for a column chart, left (vertical) for a horizontal bar
  // chart — and the VALUE axis is perpendicular to it. The previous code
  // assumed the left rule was always the value axis, so a horizontal bar
  // chart whose value axis is `<c:delete val="1">` (sample-2 slide-16) drew
  // no axis line at all even though its category axis carries an explicit
  // `<c:spPr><a:ln>`. `<a:noFill>` on a line suppresses just the rule (labels
  // stay) → `*AxisLineHidden`; an `<a:solidFill>` gives `*AxisLineColor`/Width
  // (ECMA-376 §21.2.2.* line props). Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  const catLineColor = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#aaa';
  const valLineColor = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#aaa';
  // Axis line weights from `<c:*Ax><c:spPr><a:ln@w>` (EMU). `ctx.lineWidth`
  // is in CANVAS pixels, so the point width must be multiplied by `ptToPx`
  // (px-per-point at the display scale: ~1.333 for xlsx 96dpi, ~1.05 for
  // pptx) — exactly like the chart-border code in renderChart. Without it the
  // rule is under-scaled on HiDPI/zoomed canvases. The `: 1` fallback (no
  // explicit `@w`) stays a 1px hairline.
  const catLineW = chart.catAxisLineWidthEmu ? Math.max(0.5, chart.catAxisLineWidthEmu / EMU_PER_PT) * ptToPx : 1;
  const valLineW = chart.valAxisLineWidthEmu ? Math.max(0.5, chart.valAxisLineWidthEmu / EMU_PER_PT) * ptToPx : 1;
  const strokeAxis = (x1: number, y1: number, x2: number, y2: number, color: string, lw: number): void => {
    ctx.strokeStyle = color; ctx.lineWidth = lw;
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
  };
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null;
  // Axis rules + tick marks are drawn AFTER the bars/line (see `drawAxesOnTop`
  // below) so the bars don't paint over the category baseline — PowerPoint
  // keeps the axis line crisp on top of the columns.
  const drawAxesOnTop = (): void => {
    if (!isH) {
      if (drawCatLine) strokeAxis(px0, py0 + ph, px0 + pw, py0 + ph, catLineColor, catLineW); // bottom
      if (drawValLine) strokeAxis(px0, py0, px0, py0 + ph, valLineColor, valLineW);           // left
    } else {
      if (drawCatLine) strokeAxis(px0, py0, px0, py0 + ph, catLineColor, catLineW);           // left
      if (drawValLine) strokeAxis(px0, py0 + ph, px0 + pw, py0 + ph, valLineColor, valLineW); // bottom
    }

    // Axis major tick marks (`<c:*Ax><c:majorTickMark>` — ECMA-376 §21.2.2.101).
    // PowerPoint draws short ruler ticks even when the axis rule itself is light,
    // so the bar renderer must emit them too (the line renderer already does).
    // `drawAxisTick`'s `axis` arg selects GEOMETRY: 'val' = vertical rule with
    // horizontal ticks, 'cat' = horizontal rule with vertical ticks. For a
    // column chart the value axis is vertical (left) and the category axis
    // horizontal (bottom); a horizontal bar chart swaps the two.
    if (!chart.valAxisHidden && chart.valAxisMajorTickMark && chart.valAxisMajorTickMark !== 'none') {
      for (let si = 0; si <= steps; si++) {
        const val = si * step;
        if (!isH) {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, py0 + ph - (val / axMax) * ph, valLineColor, valLineW);
        } else {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'cat', py0 + ph, px0 + (val / axMax) * pw, valLineColor, valLineW);
        }
      }
    }
    // Category ticks sit at band BOUNDARIES with crossBetween="between" (the
    // bar/column default) — the dividers between Q1|Q2|Q3|Q4 (n+1 ticks) — and
    // at category centers under "midCat".
    if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
      const onBoundary = chart.catAxisCrossBetween !== 'midCat';
      const last = onBoundary ? n : n - 1;
      for (let ci = 0; ci <= last; ci++) {
        const frac = onBoundary ? ci / n : (n === 1 ? 0.5 : ci / (n - 1));
        if (!isH) {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, px0 + frac * pw, catLineColor, catLineW);
        } else {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'val', px0, py0 + frac * ph, catLineColor, catLineW);
        }
      }
    }
  };

  // Bar cluster geometry — ECMA-376 §21.2.2.13 (gapWidth = % of bar width
  // between categories, default 150) and §21.2.2.25 (overlap = signed % of
  // bar width within a cluster, default 0). Within a cluster the pitch
  // between consecutive bars is `barW * (1 - overlap/100)`, so with N series:
  //   clusterWidth = barW + (N - 1) * barW * (1 - overlap/100)
  //   catGap       = clusterWidth + barW * gapWidth/100
  //                = barW * (1 + (N-1) * (1 - overlap/100) + gapWidth/100)
  // Solving for barW gives the formula below. Stacked charts render one bar
  // per category so we treat them as N=1 and overlap=0.
  const catGap = !isH ? pw / n : ph / n;
  const nSeriesEffective = stacked ? 1 : Math.max(1, barSeries.length);
  const overlapPct  = stacked ? 0 : (chart.barOverlap ?? 0);
  const gapWidthPct = chart.barGapWidth ?? 150;
  const denom = 1 + (nSeriesEffective - 1) * (1 - overlapPct / 100) + gapWidthPct / 100;
  const barW  = catGap / denom;
  // Pitch between bars within a cluster (not the gap — the left-edge to
  // left-edge distance). Kept named `clusterGap` for continuity with the
  // prior implementation, which also used it as a pitch.
  const clusterGap = stacked ? 0 : barW * (1 - overlapPct / 100);
  const clusterWidth = barW + (nSeriesEffective - 1) * clusterGap;
  // Center the cluster inside the category slot.
  const catStart   = (catGap - clusterWidth) / 2;

  for (let ci = 0; ci < n; ci++) {
    let stackOffset = 0;
    let stackSum = 0;
    if (pct) {
      for (const s of barSeries) stackSum += Math.abs(s.values[ci] ?? 0);
      if (stackSum === 0) stackSum = 1;
    }

    for (let si = 0; si < barSeries.length; si++) {
      const s = barSeries[si];
      const raw = s.values[ci] ?? 0;
      const val = pct ? (Math.abs(raw) / stackSum) * 100 : Math.abs(raw);
      const color = chartColor(si, s);

      if (!isH) {
        const bx = stacked
          ? px0 + ci * catGap + catStart
          : px0 + ci * catGap + catStart + si * clusterGap;
        const barH = (val / axMax) * ph;
        const by   = py0 + ph - (stacked ? (stackOffset + val) : val) / axMax * ph;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barW, barH);
        if (chart.showDataLabels && val > 0) {
          // ECMA-376 §21.2.2.30 / §21.1.2.3.10 — data label font size comes from
          // `<c:dLbls><c:txPr>...<a:defRPr@sz>` (hundredths of a point). When
          // the file specifies one we honor it; otherwise the proportional
          // heuristic keeps small bars readable.
          const lsz = chart.dataLabelFontSizeHpt
            ? (chart.dataLabelFontSizeHpt / 100) * ptToPx
            : Math.max(7, Math.min(11, barW * 0.6));
          ctx.font = `bold ${lsz}px sans-serif`;
          const text = pct
            ? `${Math.round(val)}%`
            : formatChartValWithCode(
                val,
                chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
              );
          // drawBarDataLabel takes (bx, by, barL=length, barW=thickness). For
          // a vertical column bar, "length" is the bar's height and
          // "thickness" is its horizontal width — pass them in that order.
          // Previously the args were (barW, barH) which silently swapped the
          // two and made `cx = bx + barW/2` (the horizontal-center formula
          // inside the helper) use the bar's HEIGHT instead of its width,
          // pushing data labels far to the right of the bar.
          drawBarDataLabel(
            ctx, text,
            bx, by, barH, barW,
            'vertical',
            chart.dataLabelPosition ?? null,
            s.labelColor ?? chart.dataLabelFontColor ?? null,
          );
        }
      } else {
        // Excel renders horizontal clustered bars with series 0 at the BOTTOM
        // of each category cluster (so the legend's top entry matches the bar
        // at the top of the plot). Reverse the per-series offset so `order=0`
        // ends up at the bottom; stacked horizontal bars use a single anchor.
        const siVisual = stacked ? si : (barSeries.length - 1 - si);
        const by = stacked
          ? py0 + (n - 1 - ci) * catGap + catStart
          : py0 + (n - 1 - ci) * catGap + catStart + siVisual * clusterGap;
        const barL = (val / axMax) * pw;
        const bx   = stacked ? px0 + (stackOffset / axMax) * pw : px0;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barL, barW);
        if (chart.showDataLabels && val > 0) {
          const lsz = chart.dataLabelFontSizeHpt
            ? (chart.dataLabelFontSizeHpt / 100) * ptToPx
            : Math.max(7, Math.min(11, barW * 0.6));
          ctx.font = `bold ${lsz}px sans-serif`;
          const text = pct
            ? `${Math.round(val)}%`
            : formatChartValWithCode(
                val,
                chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
              );
          drawBarDataLabel(
            ctx, text,
            bx, by, barL, barW,
            'horizontal',
            chart.dataLabelPosition ?? null,
            s.labelColor ?? chart.dataLabelFontColor ?? null,
          );
        }
      }
      if (stacked) stackOffset += val;
    }
  }

  if (!chart.catAxisHidden) {
    // `<c:catAx><c:txPr>…<a:solidFill>` colors the category tick labels (e.g.
    // sample-2 slide-16's "2025年3月期" labels are `bg1 lumMod 75%` gray).
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.font = `${Math.max(8, Math.min(11, catGap * 0.5))}px sans-serif`;
    for (let ci = 0; ci < n; ci++) {
      const label = (cats[ci] ?? '').toString().slice(0, 12);
      if (!isH) {
        const lx = px0 + ci * catGap + catGap / 2;
        ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        ctx.fillText(label, lx, py0 + ph + 3);
      } else {
        const ly = py0 + (n - 1 - ci) * catGap + catGap / 2;
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(label, px0 - 4, ly);
      }
    }
  }

  if (lineSeries.length > 0 && !isH) {
    for (let si = 0; si < lineSeries.length; si++) {
      const s = lineSeries[si];
      const color = chartColor(barSeries.length + si, s);
      // Series bound to the secondary axis map through its scale; others use
      // the primary (bar) value axis.
      const yOf = sec && s.useSecondaryAxis === true ? toYSecondary : toYPrimaryLine;
      ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.setLineDash([]);
      ctx.beginPath();
      let started = false;
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) { started = false; continue; }
        const lx = px0 + ci * catGap + catGap / 2;
        const ly = yOf(v);
        if (!started) { ctx.moveTo(lx, ly); started = true; } else ctx.lineTo(lx, ly);
      }
      ctx.stroke();
      if (s.showMarker !== false) {
        for (let ci = 0; ci < n; ci++) {
          const v = s.values[ci];
          if (v == null) continue;
          const lx = px0 + ci * catGap + catGap / 2;
          const ly = yOf(v);
          ctx.fillStyle = color;
          ctx.beginPath(); ctx.arc(lx, ly, 3, 0, Math.PI * 2); ctx.fill();
        }
      }
    }
  }

  // Primary axis rules + ticks on top of the bars/line so the category
  // baseline stays visible (the bars would otherwise paint over it).
  drawAxesOnTop();

  // Secondary value axis (right edge). Independent scale: its own "nice" major
  // unit drives the tick labels, positioned via `toYSecondary` (NOT aligned to
  // the primary gridlines — PowerPoint places them independently). Draws its
  // rule + ticks on the right; ticks mirror the left axis ("out" points right).
  if (sec) {
    const axX = px0 + pw;
    const secLineColor = sec.lineColor ? `#${sec.lineColor}` : '#aaa';
    const secLineW = sec.lineWidthEmu ? Math.max(0.5, sec.lineWidthEmu / EMU_PER_PT) * ptToPx : 1;
    if (!sec.lineHidden) strokeAxis(axX, py0, axX, py0 + ph, secLineColor, secLineW);
    if (!sec.hidden) {
      const secFontPx = sec.fontSizeHpt ? (sec.fontSizeHpt / 100) * ptToPx : secTickFontPx;
      ctx.font = `${secFontPx}px sans-serif`;
      ctx.fillStyle = sec.fontColor ? `#${sec.fontColor}` : valLabelColor;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      const tick = sec.majorTickMark;
      const tickLen = Math.max(4, secLineW + 2);
      const secSteps = Math.max(1, Math.round(sRange / sStep));
      for (let si = 0; si <= secSteps; si++) {
        const sval = sMin + si * sStep;
        const gy = toYSecondary(sval);
        if (tick && tick !== 'none') {
          const outer = tick === 'out' || tick === 'cross' ? tickLen : 0;
          const inner = tick === 'in' || tick === 'cross' ? -tickLen : 0;
          ctx.strokeStyle = secLineColor; ctx.lineWidth = secLineW;
          ctx.beginPath(); ctx.moveTo(axX + inner, gy); ctx.lineTo(axX + outer, gy); ctx.stroke();
        }
        ctx.fillText(formatChartValWithCode(sval, sec.formatCode ?? null), axX + 14, gy);
      }
    }
    if (sec.title) {
      const tFontPx = sec.titleFontSizeHpt ? (sec.titleFontSizeHpt / 100) * ptToPx : Math.max(9, h * 0.05);
      ctx.save();
      ctx.fillStyle = sec.titleFontColor
        ? `#${sec.titleFontColor}`
        : (sec.fontColor ? `#${sec.fontColor}` : '#555');
      ctx.font = `${sec.titleFontBold ? 'bold ' : ''}${tFontPx}px sans-serif`;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      // Right-axis title reads top-to-bottom (rotate +90), placed past the labels.
      ctx.translate(px0 + pw + secLabelBandW + tFontPx * 0.6, py0 + ph / 2);
      ctx.rotate(Math.PI / 2);
      ctx.fillText(sec.title, 0, 0);
      ctx.restore();
    }
  }

  // Horizontal clustered bars: Excel mirrors the series order between the
  // plot and the legend so the legend's first entry matches the top bar. We
  // already flipped the bar rendering; reverse the legend series too.
  const legendChart = isH && !stacked
    ? { ...chart, series: [...chart.series].reverse() }
    : chart;
  drawLegendForLayout(ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════

function renderLineChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  // PowerPoint's auto-layout reserves a title band with air above and below
  // the text; pinning the title to y+0 and the plot to y+titleFontPx+2 is too
  // tight. Use proportional pads so scaling preserves the same feel.
  const titleTopPad    = chart.title ? h * 0.045 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  // Axis-title bands use the real title font (XML @sz when set), independent of
  // the tick-label sizes above, so 18pt titles get a wide enough gutter.
  const { catTitlePx, valTitlePx, catTitleH, valTitleW } = axisTitleLayout(chart, w, h, ptToPx);
  // Pad based on actual label metrics rather than magic percents so an explicit
  // <c:txPr sz="1000"> (10pt) correctly compresses the plot area.
  const pad = {
    t: titleH + legTopH + valAxFontPx / 2 + 2,
    r: legRightW + w * 0.05,
    b: catAxFontPx + 12 + catTitleH + legBottomH,
    l: valAxFontPx * 2.2 + 10 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  let dataMin = Infinity; let dataMax = -Infinity;
  for (const s of chart.series) for (const v of s.values) if (v != null) { dataMin = Math.min(dataMin, v); dataMax = Math.max(dataMax, v); }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (chart.valMin != null) dataMin = chart.valMin;
  else if (dataMin > 0) dataMin = 0;
  if (chart.valMax != null) dataMax = chart.valMax;
  else if (dataMax < 0) dataMax = 0;
  if (dataMax === dataMin) dataMax = dataMin + 1;

  const { min: axMin, max: axMax, step } = valueAxisScale(dataMin, dataMax, chart.valMin, chart.valMax);
  const range = axMax - axMin; if (range === 0) return;

  const toY = (v: number) => py0 + ph - ((v - axMin) / range) * ph;
  // crossBetween="between" (default) insets the first/last category by half a
  // step so points aren't flush against the axes. "midCat" anchors them.
  const between = chart.catAxisCrossBetween !== 'midCat';
  const toX = between
    ? (i: number) => px0 + ((i + 0.5) / n) * pw
    : (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);

  if (!chart.valAxisHidden) {
    const steps = Math.round((axMax - axMin) / step);
    ctx.font = `${valAxFontPx}px sans-serif`;
    ctx.textBaseline = 'middle';
    for (let si = 0; si <= steps; si++) {
      const v = axMin + si * step;
      const gy = toY(v);
      ctx.strokeStyle = v === 0 ? '#aaa' : '#e0e0e0';
      ctx.lineWidth = v === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy);
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
      ctx.textAlign = 'right';
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode), px0 - 6, gy);
    }
  }

  // Axis lines: bottom (category) + left (value). Both default to visible
  // unless hidden explicitly. `<c:spPr><a:ln><a:noFill>` (line-only hide)
  // suppresses the rule while keeping labels and tick marks — sample-1
  // "Carbon & Growth" uses this on the value axis.
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }

  // Line width and marker size come from OOXML in points (<a:ln w=EMU> /
  // <c:marker><c:size val=pt>). We don't parse per-series overrides yet so
  // use the PowerPoint defaults (2.25pt line, 5pt marker diameter) scaled to
  // the current slide pt-per-px so both shrink with the viewport.
  const lineWidthPx = Math.max(1, 2.25 * ptToPx);
  const markerR = Math.max(2, 2.5 * ptToPx);
  const dataLabelPx = axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    ctx.strokeStyle = color; ctx.lineWidth = lineWidthPx; ctx.setLineDash([]);
    ctx.beginPath();
    let started = false;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) { started = false; continue; }
      const px = toX(ci); const py = toY(v);
      if (!started) { ctx.moveTo(px, py); started = true; } else ctx.lineTo(px, py);
    }
    ctx.stroke();
    ctx.fillStyle = color;
    // ECMA-376 §21.2.2.32 — when the series resolves to no marker, skip the
    // data-point dots but keep data labels (which pin to each raw value, not
    // to the marker).
    const drawMarkers = s.showMarker !== false;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) continue;
      if (drawMarkers) {
        ctx.beginPath(); ctx.arc(toX(ci), toY(v), markerR, 0, Math.PI * 2); ctx.fill();
      }
      if (chart.showDataLabels) {
        ctx.font = `${dataLabelPx}px sans-serif`;
        ctx.fillStyle = '#333'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        const labelOffset = drawMarkers ? markerR + 1 : 2;
        ctx.fillText(formatChartVal(v), toX(ci), toY(v) - labelOffset);
        ctx.fillStyle = color;
      }
    }
  }

  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.ceil(n / 8));
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${catAxFontPx}px sans-serif`;
    for (let ci = 0; ci < n; ci += labelInterval) {
      const tx = toX(ci);
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, tx);
      ctx.fillStyle = catLabelColor;
      ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), tx, py0 + ph + 5);
    }
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════

function renderAreaChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const stacked = chart.chartType === 'stackedArea' || chart.chartType === 'stackedAreaPct';

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const { catTitlePx, valTitlePx, catTitleH, valTitleW } = axisTitleLayout(chart, w, h, ptToPx);
  const pad = {
    t: titleH + legTopH + h * 0.02,
    r: legRightW + w * 0.05,
    b: h * 0.14 + catTitleH + legBottomH,
    l: w * 0.12 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    if (stacked) {
      let sum = 0;
      for (const s of chart.series) sum += s.values[ci] ?? 0;
      dataMax = Math.max(dataMax, sum);
    } else {
      for (const s of chart.series) dataMax = Math.max(dataMax, s.values[ci] ?? 0);
    }
  }
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;
  // Area anchors the value axis at 0; ignore the returned min.
  const { max: axMax, step } = valueAxisScale(0, dataMax, undefined, chart.valMax);

  // crossBetween="between" (Office's default; ECMA-376 §21.2.2.32 leaves the
  // default application-defined) gives each category a band of width pw/n and
  // plots its point at the band CENTER, leaving a half-band margin before the
  // first and after the last category — matching PowerPoint's Jan…Dec inset.
  // "midCat" anchors points on the category dividers (flush to the axes).
  const between = chart.catAxisCrossBetween !== 'midCat';
  const toX = between
    ? (i: number) => px0 + ((i + 0.5) / n) * pw
    : (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);
  const toY = (v: number) => py0 + ph - (v / axMax) * ph;

  // Axis line colour/weight from `<c:*Ax><c:spPr><a:ln>` (EMU → px at scale),
  // mirroring the bar/line renderers. Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  const catLineColor = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#aaa';
  const valLineColor = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#aaa';
  const catLineW = chart.catAxisLineWidthEmu ? Math.max(0.5, chart.catAxisLineWidthEmu / EMU_PER_PT) * ptToPx : 1;
  const valLineW = chart.valAxisLineWidthEmu ? Math.max(0.5, chart.valAxisLineWidthEmu / EMU_PER_PT) * ptToPx : 1;

  // Draw the translucent series fills FIRST, then lay gridlines, axis rules,
  // tick marks and labels on top so they stay visible across the filled
  // region (PowerPoint keeps the gridlines legible under the 55%-alpha area).
  const stackBase = stacked ? new Array(n).fill(0) as number[] : null;
  for (let si = chart.series.length - 1; si >= 0; si--) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    const baseY = py0 + ph;

    ctx.beginPath();
    if (stacked && stackBase) {
      for (let ci = 0; ci < n; ci++) {
        const v = (s.values[ci] ?? 0) + stackBase[ci];
        const px = toX(ci); const py = toY(v);
        if (ci === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      for (let ci = n - 1; ci >= 0; ci--) {
        ctx.lineTo(toX(ci), toY(stackBase[ci]));
      }
      for (let ci = 0; ci < n; ci++) stackBase[ci] += s.values[ci] ?? 0;
    } else {
      ctx.moveTo(toX(0), baseY);
      for (let ci = 0; ci < n; ci++) ctx.lineTo(toX(ci), toY(s.values[ci] ?? 0));
      ctx.lineTo(toX(n - 1), baseY);
    }
    ctx.closePath();
    ctx.fillStyle = hexToRgba(color, 0.6);
    ctx.fill();
    ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.setLineDash([]);
    ctx.stroke();
  }

  if (!chart.valAxisHidden) {
    ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
    ctx.textBaseline = 'middle';
    const steps = Math.round(axMax / step);
    for (let si = 0; si <= steps; si++) {
      const v = si * step; const gy = toY(v);
      ctx.strokeStyle = si === 0 ? '#aaa' : '#e0e0e0';
      ctx.lineWidth = si === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, valLineColor, valLineW);
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
      ctx.textAlign = 'right';
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode), px0 - 6, gy);
    }
  }
  // Category-axis baseline + value-axis rule. `<c:*Ax><c:spPr><a:ln><a:noFill>`
  // suppresses just the rule (labels/ticks stay) → `*AxisLineHidden`. The value
  // rule is drawn only when the file gives it a colour, matching the bar/line
  // renderers (Office's default value axis is line-less, gridlines stand in).
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.strokeStyle = catLineColor; ctx.lineWidth = catLineW;
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null) {
    ctx.strokeStyle = valLineColor; ctx.lineWidth = valLineW;
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }
  // Category-axis major tick marks. With crossBetween="between" PowerPoint
  // draws them at the band BOUNDARIES (n+1 dividers); "midCat" ticks centers.
  if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
    if (between) {
      for (let ci = 0; ci <= n; ci++) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, px0 + (ci / n) * pw, catLineColor, catLineW);
      }
    } else {
      for (let ci = 0; ci < n; ci++) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, toX(ci), catLineColor, catLineW);
      }
    }
  }

  if (!chart.catAxisHidden) {
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${Math.max(8, Math.min(11, pw / n * 0.8))}px sans-serif`;
    // Show every category label that fits; thin out only when adjacent labels
    // would collide (so 12 months all render, unlike a fixed n/8 cap).
    let maxLabelW = 0;
    for (let ci = 0; ci < n; ci++) {
      maxLabelW = Math.max(maxLabelW, ctx.measureText((cats[ci] ?? '').toString().slice(0, 10)).width);
    }
    const labelInterval = Math.max(1, Math.ceil((maxLabelW + 6) / (pw / n)));
    for (let ci = 0; ci < n; ci += labelInterval) {
      ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), toX(ci), py0 + ph + 3);
    }
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut — supports dataPointColors (per slice).
// ═══════════════════════════════════════════════════════════════════════════

function renderPieChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, isDoughnut: boolean, ptToPx: number): void {
  const { x, y, w, h } = r;
  const s = chart.series[0]; if (!s) return;
  const cats = (s.categories && s.categories.length > 0) ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  // Pie legend labels categories (one row per slice) so reserve a bit more
  // than the default 22% when placed on the side.
  const pieLeg: LegendLayout | null = chart.showLegend
    ? (() => {
        const pos = chart.legendPos ?? 'r';
        const side: LegendSide = pos === 'l' ? 'l' : pos === 't' ? 't' : pos === 'b' ? 'b' : 'r';
        if (side === 'r' || side === 'l') {
          return { side, reserveW: Math.max(80, w * 0.28), reserveH: 0 };
        }
        return { side, reserveW: 0, reserveH: Math.max(18, h * 0.08) };
      })()
    : null;
  const legRightW  = pieLeg?.side === 'r' ? pieLeg.reserveW : 0;
  const legLeftW   = pieLeg?.side === 'l' ? pieLeg.reserveW : 0;
  const legTopH    = pieLeg?.side === 't' ? pieLeg.reserveH : 0;
  const legBottomH = pieLeg?.side === 'b' ? pieLeg.reserveH : 0;

  const pw = w - legRightW - legLeftW;
  const ph = h - titleH - legTopH - legBottomH - h * 0.02;
  const cx2 = x + legLeftW + pw / 2;
  const cy2 = y + titleH + legTopH + h * 0.02 + ph / 2;
  const outerR = Math.min(pw, ph) * 0.42;
  const innerR = isDoughnut ? outerR * 0.5 : 0;

  let angle = -Math.PI / 2;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const color = pieSliceColor(i, s);
    ctx.beginPath();
    ctx.moveTo(cx2, cy2);
    ctx.arc(cx2, cy2, outerR, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = color; ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 1; ctx.stroke();

    if (chart.showDataLabels && slice > 0.15) {
      const midAngle = angle + slice / 2;
      const labelR = outerR * (isDoughnut ? 0.75 : 0.6);
      const lx2 = cx2 + Math.cos(midAngle) * labelR;
      const ly2 = cy2 + Math.sin(midAngle) * labelR;
      const pct2 = Math.round((vals[i] / total) * 100);
      const lsz = Math.max(8, outerR * 0.1);
      ctx.font = `bold ${lsz}px sans-serif`;
      ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(`${pct2}%`, lx2, ly2);
    }

    angle += slice;
  }

  if (isDoughnut) {
    ctx.beginPath(); ctx.arc(cx2, cy2, innerR, 0, Math.PI * 2);
    ctx.fillStyle = '#fff'; ctx.fill();
  }

  if (pieLeg) {
    // Pie/doughnut legends are category-driven: one row per slice, each colored
    // exactly like its slice (`pieSliceColor`). `buildLegendEntries` derives the
    // rows from the real series, so pass it through unchanged (with the resolved
    // category labels attached). The previous pseudo-series collapsed all
    // swatches to one color because it folded the series-level fill (`s.color`)
    // into every entry while the slices used the per-index palette.
    const legendSeries: ChartSeries[] = [{ ...s, categories: cats }];
    const plotLeft = cx2 - pw / 2;
    drawLegendForLayout(
      ctx, { ...chart, series: legendSeries } as ChartModel, pieLeg,
      x, y, w, h, plotLeft, cy2 - ph / 2, pw, ph, titleH + 2,
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════

function renderRadarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n < 3) return;

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH  = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const pw = w - legRightW - legLeftW;
  const ph = h - titleH - legTopH - legBottomH - h * 0.02;
  const cx2 = x + legLeftW + pw / 2;
  const cy2 = y + titleH + legTopH + h * 0.02 + ph / 2;
  const rd  = Math.min(pw, ph) * 0.38;

  let dataMax = 0;
  for (const s of chart.series) for (const v of s.values) dataMax = Math.max(dataMax, v ?? 0);
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;
  // Radar anchors the value axis at 0; ignore the returned min.
  const { max: axMax, step } = valueAxisScale(0, dataMax, undefined, chart.valMax);

  const angle0 = -Math.PI / 2;
  const spoke  = (i: number) => angle0 + (i / n) * Math.PI * 2;

  const rings = Math.round(axMax / step);
  ctx.strokeStyle = '#ddd'; ctx.lineWidth = 0.5;
  for (let ri = 1; ri <= rings; ri++) {
    const rr = (ri / rings) * rd;
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rr; const py = cy2 + Math.sin(a) * rr;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath(); ctx.stroke();
  }

  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 0.5;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    ctx.beginPath(); ctx.moveTo(cx2, cy2);
    ctx.lineTo(cx2 + Math.cos(a) * rd, cy2 + Math.sin(a) * rd); ctx.stroke();
  }

  // Radial tick labels on the top (12 o'clock) spoke — Excel places the value
  // axis there for radar charts. Respect <c:valAx><c:delete val="1"/> when the
  // caller hides the axis, and skip the 0-label at the center to avoid
  // overlapping the origin point.
  if (!chart.valAxisHidden) {
    const valAxPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
    ctx.font = `${valAxPx}px sans-serif`;
    ctx.fillStyle = '#555';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let ri = 1; ri <= rings; ri++) {
      const v = (ri / rings) * axMax;
      const rr = (ri / rings) * rd;
      ctx.fillText(formatChartVal(v), cx2 - 3, cy2 - rr);
    }
  }

  ctx.font = `${Math.max(8, Math.min(11, rd * 0.2))}px sans-serif`;
  ctx.fillStyle = '#444'; ctx.textBaseline = 'middle';
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    const lx = cx2 + Math.cos(a) * (rd + 12);
    const ly = cy2 + Math.sin(a) * (rd + 12);
    ctx.textAlign = Math.cos(a) < -0.1 ? 'right' : Math.cos(a) > 0.1 ? 'left' : 'center';
    ctx.fillText((cats[i] ?? '').toString().slice(0, 12), lx, ly);
  }

  // ECMA-376 §21.2.3.10 c:radarStyle — "filled" closes the polygon with a
  // translucent area fill; "standard" / "marker" (and default) draw the
  // line only. Markers come from per-series `<c:marker>` (which can
  // override the chart-type style by setting `<c:symbol val="none"/>`);
  // sample-1 "Biodiversity Index" sets radarStyle="marker" but every
  // series carries `<c:marker><c:symbol val="none"/>`, so Excel draws
  // lines only — no dots.
  const filled = chart.radarStyle === 'filled';
  const markerRadius = Math.max(2, rd * 0.025);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    // Build the per-spoke point list, leaving holes where the series has
    // no value (`<c:val>` ptCount > pts implies missing indices — sample-1
    // "Biodiversity Index" omits idx 0, so Excel draws an open polyline
    // from idx 1 to idx 10 without bridging back through the top spoke).
    const pts: Array<[number, number] | null> = [];
    for (let i = 0; i < n; i++) {
      const v = s.values[i];
      if (v == null) { pts.push(null); continue; }
      const frac = v / axMax;
      const a = spoke(i);
      pts.push([cx2 + Math.cos(a) * rd * frac, cy2 + Math.sin(a) * rd * frac]);
    }

    // Stroke the polyline, breaking on holes (no synthetic 0-fill).
    ctx.beginPath();
    let pen = false;
    for (const pt of pts) {
      if (pt == null) { pen = false; continue; }
      if (!pen) { ctx.moveTo(pt[0], pt[1]); pen = true; }
      else { ctx.lineTo(pt[0], pt[1]); }
    }
    // Only close the polygon when there are no gaps. With a hole anywhere
    // the radar is an open path (matches Excel's "skip missing point").
    const allPresent = pts.every(p => p != null);
    if (filled && allPresent) {
      ctx.closePath();
      ctx.fillStyle = hexToRgba(color, 0.25); ctx.fill();
    } else if (allPresent) {
      ctx.closePath();
    }
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke();

    // Markers: honor the per-series marker_symbol. When the series
    // explicitly carries `<c:marker><c:symbol val="none"/>`, the parser
    // sets showMarker=false — respect that even for radarStyle="marker"
    // charts (the chart-level style is the default; series overrides win).
    if (!filled && s.showMarker !== false) {
      ctx.fillStyle = color;
      for (const pt of pts) {
        if (pt == null) continue;
        ctx.beginPath(); ctx.arc(pt[0], pt[1], markerRadius, 0, Math.PI * 2); ctx.fill();
      }
    }
  }

  drawLegendForLayout(
    ctx, chart, leg,
    x, y, w, h,
    cx2 - pw / 2, cy2 - ph / 2, pw, ph, titleH + 2,
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart — X values from series.categories, Y from series.values.
// ═══════════════════════════════════════════════════════════════════════════

function renderScatterChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const { catTitlePx, valTitlePx, catTitleH, valTitleW } = axisTitleLayout(chart, w, h, ptToPx);

  // Title placement — manual layout overrides the auto position.
  if (chart.title) {
    const tml = chart.titleManualLayout;
    if (tml && (tml.x !== undefined || tml.y !== undefined)) {
      const tx = x + tml.x * w;
      const ty = y + tml.y * h;
      drawChartTitle(ctx, chart, tx, ty, (tml.w ?? 0.5) * w, titleFontPx);
    } else {
      drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);
    }
  }

  // Plot area placement: honor `<c:plotArea><c:manualLayout>` when present.
  // ECMA-376: layoutTarget="inner" (default) describes the inner plot rect
  // (no axes / labels); "outer" includes axes. For scatter we treat both
  // identically (the inner padding stays the same).
  const pml = chart.plotAreaManualLayout;
  let px0: number, py0: number, pw: number, ph: number;
  if (pml && pml.w != null && pml.h != null) {
    px0 = x + pml.x * w;
    py0 = y + pml.y * h;
    pw  = pml.w * w;
    ph  = pml.h * h;
  } else {
    const pad = {
      t: titleH + legTopH + h * 0.02,
      r: legRightW + w * 0.05,
      b: (chart.catAxisHidden ? h * 0.04 : h * 0.12) + catTitleH + legBottomH,
      l: (chart.valAxisHidden ? w * 0.04 : w * 0.12) + valTitleW + legLeftW,
    };
    px0 = x + pad.l; py0 = y + pad.t;
    pw = w - pad.l - pad.r; ph = h - pad.t - pad.b;
  }
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  // X / Y data extents.
  const allX: number[] = []; const allY: number[] = [];
  for (const s of chart.series) {
    const cats = s.categories ?? [];
    for (const c of cats) { const v = parseFloat(c); if (!isNaN(v)) allX.push(v); }
    for (const v of s.values) if (v != null) allY.push(v);
  }
  const useIndexX = allX.length === 0;
  if (useIndexX) {
    const maxLen = Math.max(...chart.series.map(s => s.values.length));
    for (let i = 0; i < maxLen; i++) allX.push(i);
  }

  let xMin = Math.min(...allX); let xMax = Math.max(...allX);
  let yMin = Math.min(...allY); let yMax = Math.max(...allY);
  if (xMin === xMax) { xMin -= 1; xMax += 1; }
  if (yMin === yMax) { yMin -= 1; yMax += 1; }
  // Apply explicit `<c:valAx><c:scaling><c:min/max>` and `<c:catAx>` scaling
  // when present; otherwise pad up to zero on the value axis (matches Excel
  // for charts whose data is all positive).
  if (chart.valMin != null) yMin = chart.valMin;
  else if (yMin > 0) yMin = 0;
  if (chart.valMax != null) yMax = chart.valMax;
  // Auto value (Y) axis: ONE major unit (the "nice" step) drives both the
  // rounded bounds and the gridlines — identical to bar/line/area. niceAxisMax
  // adds ~5% headroom above the data max and rounds up to that step, so the top
  // point sits below the top gridline (data 3.5 → step 0.5, max 4 → 0,.5,…,4;
  // 0.1129 → step 0.02, max 0.12). The step is taken from the DATA range and
  // reused for the gridline loop below. The post-anchor yMin (which already had
  // chart.valMin and the >0→0 anchor applied above) is the data extent; passing
  // chart.valMin/valMax as the explicit args reproduces the prior `?? niceAxis…`
  // behavior exactly. Explicit <c:valAx><c:scaling> wins. NB: the auto major
  // unit is not specified by ECMA-376 (Excel-proprietary); niceStep approximates
  // it and may differ from Excel by one step on some ranges.
  const { min: niceYMin, max: niceYMax, step: yAxisStep } =
    valueAxisScale(yMin, yMax, chart.valMin, chart.valMax);
  yMin = niceYMin; yMax = niceYMax;
  if (chart.catAxisMin != null) xMin = chart.catAxisMin;
  if (chart.catAxisMax != null) xMax = chart.catAxisMax;
  // Excel snaps auto-derived axis bounds outward to a multiple of the
  // step so both ends land on round numbers (e.g. dates jump to a date
  // before the first task and after the last). When the spec set min /
  // max explicitly we leave them alone.
  if (chart.catAxisMin == null || chart.catAxisMax == null) {
    const step = niceStep(xMax - xMin);
    if (step > 0) {
      if (chart.catAxisMin == null) xMin = Math.floor(xMin / step) * step;
      if (chart.catAxisMax == null) xMax = Math.ceil(xMax / step) * step;
    }
  }

  const toX = (v: number) => px0 + ((v - xMin) / (xMax - xMin)) * pw;
  const toY = (v: number) => py0 + ph - ((v - yMin) / (yMax - yMin)) * ph;

  // Y-axis gridlines + labels + major tick marks.
  if (!chart.valAxisHidden) {
    const yTickFontPx = Math.max(8, Math.min(11, ph / 20));
    ctx.font = `${chart.valAxisFontBold ? 'bold ' : ''}${yTickFontPx}px sans-serif`;
    const ySteps = Math.round((yMax - yMin) / yAxisStep) + 1;
    for (let si = 0; si < ySteps; si++) {
      const v = yMin + si * yAxisStep; if (v > yMax + yAxisStep * 0.01) break;
      const gy = toY(v);
      ctx.strokeStyle = '#e0e0e0'; ctx.lineWidth = 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillStyle = '#555'; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode), px0 - 4, gy);
      const yAxisLineColor = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : undefined;
      const yAxisLineWidth = chart.valAxisLineWidthEmu
        ? Math.max(0.5, chart.valAxisLineWidthEmu / EMU_PER_PT) * ptToPx : undefined;
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, yAxisLineColor, yAxisLineWidth);
    }
  }

  // X-axis Y position. `<c:catAx><c:crossesAt>` wins; otherwise honor
  // `<c:catAx><c:crosses>` (`autoZero` / `min` / `max`). The `autoZero`
  // default puts the axis at y=0 if the data range crosses zero —
  // that's what makes Excel "Project Timeline" templates split
  // milestones (positive Y) above the timeline ruler and tasks
  // (negative Y) below. Clamped to the plot rect so the axis line
  // stays visible when the data doesn't actually cross.
  let xAxisY = py0 + ph;
  if (chart.catAxisCrossesAt != null) {
    xAxisY = clamp(toY(chart.catAxisCrossesAt), py0, py0 + ph);
  } else {
    const c = chart.catAxisCrosses ?? 'autoZero';
    if (c === 'autoZero' && yMin < 0 && yMax > 0) {
      xAxisY = clamp(toY(0), py0, py0 + ph);
    } else if (c === 'min') {
      xAxisY = py0 + ph;
    } else if (c === 'max') {
      xAxisY = py0;
    }
  }

  // X-axis line (the timeline ruler in Gantt-style scatter charts depends
  // on this line's stroke). Tick labels are skipped when the category axis
  // is hidden via `<c:delete val="1"/>`; the rule itself is also gated on
  // `<c:catAx><c:spPr><a:ln><a:noFill>` (line-only hide). Color and
  // weight come from `<c:catAx><c:spPr><a:ln>` when present; default
  // otherwise.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.save();
    ctx.strokeStyle = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#888';
    ctx.lineWidth = chart.catAxisLineWidthEmu
      ? Math.max(0.5, chart.catAxisLineWidthEmu / EMU_PER_PT) * ptToPx
      : 1;
    ctx.lineCap = 'butt';
    ctx.beginPath(); ctx.moveTo(px0, xAxisY); ctx.lineTo(px0 + pw, xAxisY); ctx.stroke();
    ctx.restore();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.save();
    ctx.strokeStyle = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#888';
    ctx.lineWidth = chart.valAxisLineWidthEmu
      ? Math.max(0.5, chart.valAxisLineWidthEmu / EMU_PER_PT) * ptToPx
      : 1;
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
    ctx.restore();
  }

  // X-axis tick labels (catAxis), formatted via catAxisFormatCode (typically
  // a date code like "m/d/yyyy"). Skipped when catAxisHidden. Drawn just
  // below the axis line wherever it sits (axis crossing in the middle of
  // the plot still anchors the labels to the line itself). Major tick
  // marks also get drawn here so `<c:majorTickMark val="cross">` produces
  // the crossing ruler look that templates like the Vertex42 timeline
  // depend on.
  if (!chart.catAxisHidden) {
    const tickFontPx = Math.max(8, Math.min(11, ph / 20));
    ctx.font = `${chart.catAxisFontBold ? 'bold ' : ''}${tickFontPx}px sans-serif`;
    const xStep = niceStep(xMax - xMin);
    const xSteps = Math.round((xMax - xMin) / xStep) + 1;
    ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    for (let si = 0; si < xSteps; si++) {
      const v = xMin + si * xStep; if (v > xMax + xStep * 0.01) break;
      const gx = toX(v);
      ctx.fillText(formatChartValWithCode(v, chart.catAxisFormatCode), gx, xAxisY + 4);
      const xAxisLineColor = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : undefined;
      const xAxisLineWidth = chart.catAxisLineWidthEmu
        ? Math.max(0.5, chart.catAxisLineWidthEmu / EMU_PER_PT) * ptToPx : undefined;
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', xAxisY, gx, xAxisLineColor, xAxisLineWidth);
    }
  }

  // ECMA-376 §21.2.2.42 `<c:scatterStyle>`. Drives whether scatter points
  // are connected (line / smooth) and whether markers are also drawn.
  // For bubble charts the value is ignored (always markers, sized by data).
  const isBubble = chart.chartType === 'bubble';
  const style = isBubble ? 'marker' : (chart.scatterStyle ?? 'marker');
  const drawLines     = style === 'line' || style === 'lineMarker' || style === 'lineNoMarker';
  const drawSmooth    = style === 'smooth' || style === 'smoothMarker' || style === 'smoothNoMarker';
  const hideMarkersByStyle = style === 'lineNoMarker' || style === 'smoothNoMarker';

  // Render each series. Order: error bars (behind), connecting lines,
  // markers, then data labels (in front). dPt overrides apply per point
  // for color and marker shape; dLbl overrides apply per point for label
  // text and position.
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const fallbackColor = chartColor(si, s);
    const cats = s.categories ?? [];

    // Error bars (drawn first so markers overlay the bar tip).
    for (const eb of s.errBars ?? []) {
      drawSeriesErrorBars(ctx, s, eb, cats, useIndexX, toX, toY, fallbackColor);
    }

    // Connecting lines (scatterStyle = line / smooth / lineMarker / smoothMarker).
    if (drawLines || drawSmooth) {
      const pts: Array<{ x: number; y: number }> = [];
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci]; if (yv == null) continue;
        const xv = useIndexX ? ci : parseFloat(cats[ci] ?? '0');
        if (isNaN(xv)) continue;
        pts.push({ x: toX(xv), y: toY(yv) });
      }
      if (pts.length >= 2) {
        ctx.save();
        ctx.strokeStyle = s.color ? `#${s.color}` : fallbackColor;
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        if (drawSmooth && pts.length >= 3) {
          // Catmull-Rom-ish: cubic Bézier between consecutive points with
          // tangents derived from neighbours. Good enough for the typical
          // ECMA-376 smoothing intent without shipping a full spline lib.
          for (let i = 0; i < pts.length - 1; i++) {
            const p0 = pts[i - 1] ?? pts[i];
            const p1 = pts[i];
            const p2 = pts[i + 1];
            const p3 = pts[i + 2] ?? p2;
            const cp1x = p1.x + (p2.x - p0.x) / 6;
            const cp1y = p1.y + (p2.y - p0.y) / 6;
            const cp2x = p2.x - (p3.x - p1.x) / 6;
            const cp2y = p2.y - (p3.y - p1.y) / 6;
            ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
          }
        } else {
          for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
        }
        ctx.stroke();
        ctx.restore();
      }
    }

    // Markers (skip when symbol="none", series-level showMarker false, or
    // the scatter style explicitly disables markers).
    const hideMarkers = hideMarkersByStyle
      || s.showMarker === false
      || (typeof s.markerSymbol === 'string' && s.markerSymbol === 'none');
    if (!hideMarkers) {
      // Bubble size scaling. ECMA-376 §21.2.2.4 treats `<c:bubbleSize>` as an
      // area-proportional value, so radius scales by sqrt. We pick a max
      // radius proportional to the plot width / point count so bubbles
      // don't overlap in typical Excel-style data.
      let bubbleScale = 0;
      if (isBubble && s.bubbleSizes && s.bubbleSizes.length > 0) {
        const maxSz = Math.max(0, ...s.bubbleSizes.filter((v): v is number => v != null));
        if (maxSz > 0) {
          const maxRadiusPx = Math.min(pw, ph) / Math.max(8, s.values.length * 1.6);
          bubbleScale = maxRadiusPx / Math.sqrt(maxSz);
        }
      }

      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci]; if (yv == null) continue;
        const xv = useIndexX ? ci : parseFloat(cats[ci] ?? '0');
        if (isNaN(xv)) continue;
        const dpt = (s.dataPointOverrides ?? []).find(d => d.idx === ci);
        const symbol = (dpt?.markerSymbol ?? s.markerSymbol ?? 'circle') as string;
        let sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
        if (isBubble && bubbleScale > 0) {
          const bsz = s.bubbleSizes?.[ci];
          if (bsz != null && bsz > 0) {
            // Convert resulting radius (px) back to pt so drawMarker's
            // ptToPx multiplication gives the same px size.
            sizePt = (Math.sqrt(bsz) * bubbleScale * 2) / ptToPx;
          }
        }
        const fill = dpt?.markerFill
          ?? dpt?.color
          ?? s.markerFill
          ?? fallbackColor;
        const line = dpt?.markerLine ?? s.markerLine ?? null;
        drawMarker(ctx, toX(xv), toY(yv), symbol, sizePt, fill, line, ptToPx);
      }
    }

    // Per-point data labels (`<c:dLbl idx>`) and series-level defaults.
    drawSeriesDataLabels(ctx, s, cats, useIndexX, toX, toY, ph, ptToPx);
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}

/** Draw a single ECMA-376 §21.2.2.32 marker shape centered at `(cx, cy)`.
 *  `sizePt` is the spec's marker side length in points (Excel's default
 *  is 5). `fill` and `line` are hex strings; a leading `#` is tolerated so
 *  callers that route through `chartColor` (which returns `#RRGGBB`)
 *  don't end up double-prefixing into an invalid `##RRGGBB`. `line` may
 *  be null in which case no outline is drawn. `picture` falls back to a
 *  square because we don't ship the embedded image yet. */
function drawMarker(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  symbol: string,
  sizePt: number,
  fill: string,
  line: string | null,
  ptToPx: number,
): void {
  const sizePx = Math.max(2, sizePt * ptToPx);
  const half = sizePx / 2;
  const fillCss = fill.startsWith('#') ? fill : `#${fill}`;
  const lineCss = line ? (line.startsWith('#') ? line : `#${line}`) : null;
  ctx.save();
  ctx.fillStyle = fillCss;
  if (lineCss) {
    ctx.strokeStyle = lineCss;
    ctx.lineWidth = 1;
  }
  switch (symbol) {
    case 'square': {
      ctx.fillRect(cx - half, cy - half, sizePx, sizePx);
      if (line) ctx.strokeRect(cx - half, cy - half, sizePx, sizePx);
      break;
    }
    case 'diamond': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy);
      ctx.lineTo(cx, cy + half);
      ctx.lineTo(cx - half, cy);
      ctx.closePath();
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
    case 'triangle': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy + half);
      ctx.lineTo(cx - half, cy + half);
      ctx.closePath();
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
    case 'x': {
      ctx.strokeStyle = fillCss;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy - half); ctx.lineTo(cx + half, cy + half);
      ctx.moveTo(cx - half, cy + half); ctx.lineTo(cx + half, cy - half);
      ctx.stroke();
      break;
    }
    case 'plus': {
      ctx.strokeStyle = fillCss;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy); ctx.lineTo(cx + half, cy);
      ctx.moveTo(cx, cy - half); ctx.lineTo(cx, cy + half);
      ctx.stroke();
      break;
    }
    case 'star': {
      // 5-point star inscribed in a circle of radius `half`.
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? half : half * 0.45;
        const a = -Math.PI / 2 + i * Math.PI / 5;
        const px = cx + Math.cos(a) * r;
        const py = cy + Math.sin(a) * r;
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.closePath();
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
    case 'dot': {
      // Excel's "dot" is a small filled circle ~half the size of "circle".
      ctx.beginPath(); ctx.arc(cx, cy, Math.max(1, sizePx * 0.25), 0, Math.PI * 2); ctx.fill();
      break;
    }
    case 'dash': {
      const dh = Math.max(1, sizePx * 0.25);
      ctx.fillRect(cx - half, cy - dh / 2, sizePx, dh);
      break;
    }
    case 'picture':
    case 'circle':
    default: {
      ctx.beginPath();
      ctx.arc(cx, cy, half, 0, Math.PI * 2);
      ctx.fill();
      if (line) ctx.stroke();
      break;
    }
  }
  ctx.restore();
}

/** Draw error bars for one series + one direction. Each segment is a line
 *  from the data point to the offset point, plus an optional perpendicular
 *  end-cap (skipped when `eb.noEndCap`). */
function drawSeriesErrorBars(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  eb: NonNullable<ChartSeries['errBars']>[number],
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  fallbackColor: string,
): void {
  ctx.save();
  ctx.strokeStyle = eb.color ? `#${eb.color}` : fallbackColor;
  ctx.lineWidth = eb.lineWidthEmu ? Math.max(0.5, eb.lineWidthEmu / EMU_PER_PT) : 1;
  ctx.setLineDash(dashPatternForPreset(eb.dash));
  const drawPlus = eb.barType === 'plus' || eb.barType === 'both';
  const drawMinus = eb.barType === 'minus' || eb.barType === 'both';
  const isX = eb.dir === 'x';
  const capHalf = ctx.lineWidth * 1.5;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    const xv = useIndexX ? i : parseFloat(cats[i] ?? '0');
    if (isNaN(xv)) continue;
    const px = toX(xv); const py = toY(yv);
    const drawSeg = (dataDelta: number) => {
      let x2 = px, y2 = py;
      if (isX) {
        // X delta is in data X units, so map (xv + delta) → px. For the
        // minus side delta is already a positive magnitude, flip the sign.
        x2 = toX(xv + dataDelta);
      } else {
        // Y delta similar; positive moves the bar toward higher data values
        // (visually upward for our orientation).
        y2 = toY(yv + dataDelta);
      }
      ctx.beginPath();
      ctx.moveTo(px, py); ctx.lineTo(x2, y2); ctx.stroke();
      if (!eb.noEndCap) {
        ctx.save(); ctx.setLineDash([]);
        ctx.beginPath();
        if (isX) {
          ctx.moveTo(x2, y2 - capHalf); ctx.lineTo(x2, y2 + capHalf);
        } else {
          ctx.moveTo(x2 - capHalf, y2); ctx.lineTo(x2 + capHalf, y2);
        }
        ctx.stroke();
        ctx.restore();
      }
    };
    // ECMA-376 §21.2.2.20: plus side is `point + plus[i]`, minus side is
    // `point - minus[i]`. For `cust` errValType the values may be signed
    // (e.g. negative minus values that effectively flip direction); for
    // `fixedVal`/`stdErr`/`stdDev`/`percentage` the parser stores positive
    // magnitudes, so the same formula gives the expected direction.
    if (drawPlus) {
      const v = eb.plus[i]; if (v != null) drawSeg(v);
    }
    if (drawMinus) {
      const v = eb.minus[i]; if (v != null) drawSeg(-v);
    }
  }
  ctx.restore();
}

/** Draw per-point data labels: position-aware text near each marker. */
function drawSeriesDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  ph: number,
  ptToPx: number,
): void {
  const overrides = s.dataLabelOverrides ?? [];
  if (overrides.length === 0 && !s.seriesDataLabels) return;
  const seriesDef = s.seriesDataLabels;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    const xv = useIndexX ? i : parseFloat(cats[i] ?? '0');
    if (isNaN(xv)) continue;
    const ovr = overrides.find(o => o.idx === i);
    let text: string;
    if (ovr) {
      // `<c:delete val="1"/>` produced an empty text — skip drawing.
      if (ovr.text === '') continue;
      text = ovr.text;
    } else if (seriesDef && (seriesDef.showVal || seriesDef.showSerName || seriesDef.showCatName)) {
      const parts: string[] = [];
      if (seriesDef.showCatName && !useIndexX) parts.push(cats[i] ?? '');
      if (seriesDef.showSerName) parts.push(s.name);
      if (seriesDef.showVal) {
        parts.push(formatChartValWithCode(yv, seriesDef.formatCode ?? null));
      }
      text = parts.filter(Boolean).join(' ');
      if (!text) continue;
    } else {
      continue;
    }
    const pos = ovr?.position ?? seriesDef?.position ?? 'r';
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = sizeHpt
      ? (sizeHpt / 100) * ptToPx
      : Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    drawDataLabelText(ctx, toX(xv), toY(yv), text, pos, fontSizePx, color, bold);
  }
}

function drawDataLabelText(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  text: string,
  position: string,
  fontSizePx: number,
  color: string | undefined,
  bold: boolean,
): void {
  ctx.save();
  ctx.font = `${bold ? 'bold ' : ''}${fontSizePx}px sans-serif`;
  ctx.fillStyle = color ? `#${color}` : '#333';
  const offset = fontSizePx * 0.6;
  let tx = cx, ty = cy;
  switch (position) {
    case 'l':
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      tx = cx - offset; break;
    case 'r':
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      tx = cx + offset; break;
    case 't':
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      ty = cy - offset; break;
    case 'b':
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ty = cy + offset; break;
    case 'ctr':
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      break;
    default:
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      tx = cx + offset; break;
  }
  // Multi-line labels: split on newline and stack vertically.
  const lines = text.split(/\r?\n/);
  const lineH = fontSizePx * 1.15;
  const totalH = lineH * lines.length;
  let lineY = ty;
  if (ctx.textBaseline === 'middle') lineY = ty - (totalH - lineH) / 2;
  else if (ctx.textBaseline === 'bottom') lineY = ty - (totalH - lineH);
  for (const line of lines) {
    ctx.fillText(line, tx, lineY);
    lineY += lineH;
  }
  ctx.restore();
}

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

function dashPatternForPreset(preset: string | undefined): number[] {
  if (!preset) return [];
  switch (preset) {
    case 'solid':                  return [];
    case 'dot':       case 'sysDot': return [1, 2];
    case 'dash':      case 'sysDash':return [4, 2];
    case 'lgDash':                  return [8, 3];
    case 'dashDot':   case 'sysDashDot':   return [4, 2, 1, 2];
    case 'lgDashDot':                       return [8, 3, 1, 3];
    case 'dashDotDot':case 'sysDashDotDot':case 'lgDashDotDot': return [4, 2, 1, 2, 1, 2];
    default: return [];
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Waterfall chart — subtotal bars filled, delta bars outlined.
// ═══════════════════════════════════════════════════════════════════════════

function renderWaterfallChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  // PowerPoint's chartEx waterfall uses very thin side margins when the
  // value axis is hidden — there's no axis label area to reserve.
  // An explicit plot-area layout wins, exactly as the other chart renderers
  // do (ECMA-376 §21.2.2.32 `<c:plotArea><c:layout><c:manualLayout>`).
  const pml = chart.plotAreaManualLayout;
  let px0: number, py0: number, pw: number, ph: number;
  if (pml && pml.w != null && pml.h != null) {
    px0 = x + pml.x * w;
    py0 = y + pml.y * h;
    pw  = pml.w * w;
    ph  = pml.h * h;
  } else {
    // Fallback plot insets when the chart gives no explicit layout. The side
    // margins are principled (thin when the value axis is hidden, otherwise a
    // value-label gutter). The top/bottom values are HEURISTIC: 0.12 / 0.14
    // were tuned against sample-2 slide-8 (its callout tips assume a specific
    // running-total line y) and are NOT a documented chartEx default. Replace
    // once chartEx `<cx:plotArea><cx:layout>` is parsed or PowerPoint's default
    // waterfall insets are confirmed. Tracked — do not extend this tuning.
    const padL = chart.valAxisHidden ? w * 0.01 : w * 0.11;
    const padR = w * 0.01;
    const padT = h * 0.12;
    const padB = h * 0.14;
    px0 = x + padL;
    py0 = y + padT;
    pw  = w - padL - padR;
    ph  = h - padT - padB;
  }

  const vals = chart.series[0]?.values ?? [];
  const cats = chart.categories;
  const n = cats.length;
  if (n === 0) return;

  const subSet = new Set(chart.subtotalIndices);

  let running = 0;
  const bars: Array<{ start: number; end: number; isSub: boolean; isPos: boolean }> = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i] ?? 0;
    const isSub = i === 0 || subSet.has(i);
    if (isSub) {
      bars.push({ start: 0, end: v, isSub: true, isPos: true });
      running = v;
    } else {
      const start = v >= 0 ? running : running + v;
      const end   = v >= 0 ? running + v : running;
      bars.push({ start, end, isSub: false, isPos: v >= 0 });
      running += v;
    }
  }

  const allEnds = bars.map(b => b.end);
  const allStarts = bars.map(b => b.start);
  const rawMax = Math.max(...allEnds, ...allStarts);
  const rawMin = Math.min(...allStarts, 0);
  const dataRange = rawMax - rawMin;
  if (dataRange <= 0) return;
  // PowerPoint anchors waterfall bars to value=0 when all data is non-negative
  // (the x-axis sits flush against the bar bases). Adding a 5% pad below 0
  // would lift the bars off the axis. Only pad when there are actual negatives
  // to display below zero.
  const dataMin = rawMin < 0 ? rawMin - dataRange * 0.05 : 0;
  const padded = (rawMax - dataMin) * 1.1;
  const dataMax = dataMin + padded;

  const step = niceStep(padded);
  ctx.save();
  const fontSize = Math.round(h * 0.042);
  ctx.font = `${fontSize}px sans-serif`;

  // ECMA-376 / chartEx §axis@hidden: when the value axis is hidden, skip the
  // value-axis gridlines, tick labels and the left segment of the L-frame.
  // This is the canonical PowerPoint look for waterfall analyses where the
  // value scale is implicit in the data labels on each bar.
  if (!chart.valAxisHidden) {
    ctx.strokeStyle = '#e8e8e8';
    ctx.lineWidth = 0.7;
    ctx.fillStyle = '#666';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let v = Math.ceil(dataMin / step) * step; v <= dataMax; v += step) {
      const gy = py0 + ph * (1 - (v - dataMin) / padded);
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillText(v.toLocaleString(), px0 - 4, gy);
    }
  }

  // L-frame: vertical (value-axis) rule + horizontal (category-axis) baseline.
  // Each segment is independently gated on its axis's `<c:delete>` *and*
  // `<c:spPr><a:ln><a:noFill>` (line-only hide).
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden;
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (drawValLine || drawCatLine) {
    ctx.strokeStyle = '#bbb';
    ctx.lineWidth = 1;
    ctx.beginPath();
    if (drawValLine) {
      ctx.moveTo(px0, py0);
      ctx.lineTo(px0, py0 + ph);
      if (drawCatLine) ctx.lineTo(px0 + pw, py0 + ph);
    } else if (drawCatLine) {
      ctx.moveTo(px0, py0 + ph);
      ctx.lineTo(px0 + pw, py0 + ph);
    }
    ctx.stroke();
  }

  const colorSub = '#196ECA';
  const colorPos = '#5BA4E6';
  const colorNeg = '#E46970';

  // ECMA-376 / chartEx §17.18.34 ST_GapAmount: gapWidth is the gap between
  // adjacent categories expressed as a percentage of the bar width
  // (legacy `<c:gapWidth val>`) or as a fraction (chartEx
  // `<cx:catScaling gapWidth>`, normalised to the same percent form by the
  // parser). The bar then occupies `catGap / (1 + gapWidth/100)`. Default
  // 150% per the spec when neither attribute is present.
  const gapW = pw / n;
  const gapWidthPct = chart.barGapWidth ?? 150;
  const barW = gapW / (1 + gapWidthPct / 100);

  bars.forEach((bar, i) => {
    const bx = px0 + gapW * i + (gapW - barW) / 2;
    const yTop = py0 + ph * (1 - (bar.end - dataMin) / padded);
    const yBot = py0 + ph * (1 - (bar.start - dataMin) / padded);
    const bh = Math.max(1, yBot - yTop);

    if (bar.isSub) {
      ctx.fillStyle = colorSub;
      ctx.fillRect(bx, yTop, barW, bh);
    } else {
      ctx.strokeStyle = bar.isPos ? colorPos : colorNeg;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(bx + 0.75, yTop + 0.75, barW - 1.5, bh - 1.5);
    }

    if (i < n - 1) {
      const nextBx = px0 + gapW * (i + 1) + (gapW - barW) / 2;
      const connY = bar.isPos ? yTop : yBot;
      ctx.strokeStyle = '#ccc';
      ctx.lineWidth = 0.8;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(bx + barW, connY);
      ctx.lineTo(nextBx, connY);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    const rawVal = vals[i] ?? 0;
    const labelText = rawVal < 0
      ? `△ ${Math.abs(rawVal).toLocaleString()}`
      : rawVal.toLocaleString();
    // Per-data-point label colour from chartEx `<cx:dataLabel idx>` (parsed
    // into series.dataLabelColors). Falls back to chart.dataLabelFontColor,
    // then to neutral grey. PowerPoint paints negative-bar labels in
    // accent1 (red) for sample-2's waterfall.
    const perPointColor = chart.series[0]?.dataLabelColors?.[i] ?? null;
    const labelColor = perPointColor
      ? `#${perPointColor}`
      : chart.dataLabelFontColor
        ? `#${chart.dataLabelFontColor}`
        : '#595959';
    ctx.fillStyle = labelColor;
    ctx.font = `bold ${Math.round(h * 0.044)}px sans-serif`;
    ctx.textAlign = 'center';
    // Negative bars: label sits BELOW the bar (`outEnd` for a negative value
    // points downward in chartEx). Positive bars and subtotals: label ABOVE.
    if (rawVal < 0) {
      ctx.textBaseline = 'top';
      ctx.fillText(labelText, bx + barW / 2, yBot + 3);
    } else {
      ctx.textBaseline = 'bottom';
      ctx.fillText(labelText, bx + barW / 2, yTop - 3);
    }
  });

  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#666';
  ctx.font = `${Math.round(h * 0.038)}px sans-serif`;
  const labelY = py0 + ph + 4;
  for (let i = 0; i < n; i++) {
    const ccx = px0 + gapW * i + gapW / 2;
    const lines = cats[i].split(/\s+/);
    lines.forEach((line, li) => ctx.fillText(line, ccx, labelY + li * (fontSize + 2)));
  }

  ctx.restore();
}

// ─── Background frame + dispatcher ──────────────────────────────────────────

/**
 * Render a chart (background frame + dispatch on `chartType`).
 * `rect` is in pixel coordinates on the target canvas.
 */
export function renderChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  /**
   * Pixels per point at the caller's current display scale. For PPTX at
   * 960px/12192000EMU the value is ~1.05; xlsx's sheet view renders at
   * device-px where 1pt≈1.333. Used to size title/axis labels whose
   * XML-specified sizes are in OOXML hundredths of a point.
   */
  ptToPx: number = PT_TO_PX,
): void {
  const { x, y, w, h } = rect;
  // Only fill the outer chartSpace when chartBg is set; a null means noFill
  // (transparent) per OOXML, so the underlying slide/sheet shows through.
  if (chart.chartBg) {
    ctx.fillStyle = `#${chart.chartBg}`;
    ctx.fillRect(x, y, w, h);
  }

  // Explicit chart border — drawn ONLY when the XML declared a paintable
  // `<c:chartSpace><c:spPr><a:ln><a:solidFill>` (chartBorderColor is null
  // otherwise; there is no default Excel-style frame). Width comes from
  // `<a:ln@w>` (EMU → pt → px); absent width falls back to a 1px hairline.
  if (chart.chartBorderColor) {
    ctx.save();
    ctx.strokeStyle = `#${chart.chartBorderColor}`;
    // `<a:ln>` with no `@w` means width 0 per ECMA-376 §20.1.2.2.24, i.e. invisible;
    // but Excel renders a fill-without-width line as a ~hairline, so we draw 1px to
    // match the app rather than dropping a declared border.
    ctx.lineWidth = chart.chartBorderWidthEmu
      ? Math.max(0.5, chart.chartBorderWidthEmu / EMU_PER_PT) * ptToPx
      : 1;
    // Inset by half the line width so the full stroke stays inside the rect.
    const lw = ctx.lineWidth;
    ctx.strokeRect(x + lw / 2, y + lw / 2, w - lw, h - lw);
    ctx.restore();
  }

  if (chart.series.length === 0) {
    ctx.fillStyle = '#888';
    ctx.font = '12px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('(no data)', x + w / 2, y + h / 2);
    return;
  }

  switch (chart.chartType) {
    case 'clusteredBar':
    case 'clusteredBarH':
    case 'stackedBar':
    case 'stackedBarH':
    case 'stackedBarPct':
    case 'stackedBarHPct':
      renderBarChart(ctx, chart, rect, ptToPx); break;
    case 'line':
    case 'stackedLine':
    case 'stackedLinePct':
      renderLineChart(ctx, chart, rect, ptToPx); break;
    case 'area':
    case 'stackedArea':
    case 'stackedAreaPct':
      renderAreaChart(ctx, chart, rect, ptToPx); break;
    case 'pie':
      renderPieChart(ctx, chart, rect, false, ptToPx); break;
    case 'doughnut':
      renderPieChart(ctx, chart, rect, true, ptToPx); break;
    case 'radar':
      renderRadarChart(ctx, chart, rect, ptToPx); break;
    case 'scatter':
    case 'bubble':
      renderScatterChart(ctx, chart, rect, ptToPx); break;
    case 'waterfall':
      renderWaterfallChart(ctx, chart, rect); break;
    default:
      ctx.fillStyle = '#888';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`Chart: ${chart.chartType}`, x + w / 2, y + h / 2);
  }
}
