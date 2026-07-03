// ===== Unified chart model =====
// Shared by @silurus/ooxml-pptx and @silurus/ooxml-xlsx.
//
// Parser JSON from each format is adapted into `ChartModel` and then passed
// to `renderChart` in @silurus/ooxml-core. This keeps a single source of
// truth for chart rendering across PPTX / XLSX (and future DrawingML charts
// in DOCX).

export interface ChartSeries {
  name: string;
  /** Hex without '#'. null = fall back to palette. */
  color: string | null;
  /** Numeric values; null = missing data point. */
  values: (number | null)[];
  /**
   * Per-data-point colors (pie / doughnut). Hex without '#'. null inside the
   * array = use palette for that slice. Omit entirely for non-pie series.
   */
  dataPointColors?: (string | null)[] | null;
  /**
   * Per-data-point data-label text colors. Used by chartEx (`<cx:dataLabel idx>`)
   * to override label colour per bar — sample-2's waterfall paints negative
   * △ values in red while positive values stay black. Null inside the array =
   * fall back to the chart-level `dataLabelFontColor`.
   */
  dataLabelColors?: (string | null)[] | null;
  /**
   * Series-level data-label text colour (`<c:ser><c:dLbls><c:txPr>…solidFill`,
   * ECMA-376 §21.2.2.216). Hex without '#'. Stacked-bar charts colour each
   * segment's label independently (e.g. white on the dark segment, black on
   * the light one), which a single chart-level `dataLabelFontColor` can't
   * express. Takes precedence over `dataLabelFontColor`; null = no override.
   */
  labelColor?: string | null;
  /**
   * Mixed chart: per-series chart type override. Currently only "line" (XLSX
   * and PPTX combo charts) is honoured; other values are treated as the
   * chart's primary type.
   */
  seriesType?: string | null;
  /**
   * Combo chart: this series is plotted against the SECONDARY value axis
   * (`ChartModel.secondaryValAxis`) — the `<c:valAx>` with `axPos="r"` /
   * `<c:crosses val="max">`. When false/absent the series uses the primary
   * (left) value-axis scale. PowerPoint's "Revenue vs. gross margin" combo
   * (sample-14 slide-8) puts the margin line on a 0–100% secondary axis.
   */
  useSecondaryAxis?: boolean | null;
  /**
   * Scatter-only X values (as strings). When null the series uses
   * `ChartModel.categories` as X.
   */
  categories?: string[] | null;
  /**
   * Resolved marker visibility for line/scatter series. ECMA-376 §21.2.2.32
   * `<c:marker><c:symbol>` defaults to "none" for line charts unless the
   * chart-level `<c:marker val="1"/>` or a per-series symbol opts in. When
   * undefined/null the renderer uses its own default (visible) so callers
   * that don't parse markers (e.g. pptx today) keep their existing behavior.
   */
  showMarker?: boolean | null;
  /**
   * Excel number-format code for this series' values (ECMA-376 §21.2.2.37,
   * `<c:val>/<c:numRef>/<c:formatCode>`). Used to format data labels when the
   * chart-level `<c:dLbls><c:numFmt>` is not set. null = no series-level code.
   */
  valFormatCode?: string | null;
  /**
   * `<c:marker><c:symbol val>` (ECMA-376 §21.2.2.32) — point marker shape.
   * One of "circle"|"square"|"diamond"|"triangle"|"x"|"plus"|"star"|
   * "dot"|"dash"|"picture"|"none". null = renderer default (circle when
   * showMarker is true).
   */
  markerSymbol?: string | null;
  /**
   * `<c:marker><c:size val>` (ECMA-376 §21.2.2.34) — marker side length in
   * points. null = renderer default (~5 pt).
   */
  markerSize?: number | null;
  /** `<c:marker><c:spPr><a:solidFill>` resolved hex (no `#`). */
  markerFill?: string | null;
  /** `<c:marker><c:spPr><a:ln><a:solidFill>` resolved hex (no `#`). */
  markerLine?: string | null;
  /**
   * Per-data-point overrides (ECMA-376 §21.2.2.39 `<c:dPt>`). Keyed by point
   * index. Any unset field falls back to the series-level value.
   */
  dataPointOverrides?: ChartDataPointOverride[] | null;
  /**
   * Per-data-point custom labels (ECMA-376 §21.2.2.45 `<c:dLbl idx>`).
   * `text` is the resolved plain string — `<a:fld type="CELLRANGE">`
   * placeholders are already substituted at parse time. An empty string
   * means the point's label was deleted with `<c:delete val="1"/>` and
   * the renderer should skip it.
   */
  dataLabelOverrides?: ChartDataLabelOverride[] | null;
  /**
   * Series-level `<c:dLbls>` block (showVal / showSerName / position).
   * Applied to every point lacking its own `<c:dLbl>` override.
   */
  seriesDataLabels?: ChartSeriesDataLabels | null;
  /**
   * `<c:errBars>` per-series error bars (ECMA-376 §21.2.2.20). Up to two
   * (one per direction). Plus / minus deltas are absolute per-point values
   * regardless of `errValType`.
   */
  errBars?: ChartErrBars[] | null;
  /**
   * `<c:bubbleSize>` per-point sizes for bubble charts (ECMA-376 §21.2.2.4).
   * Drives marker radius — renderer treats the values as areas (radius
   * scales by sqrt) so visual area is proportional to value, matching
   * Excel. null / empty array = uniform marker size. Ignored for non-bubble
   * series.
   */
  bubbleSizes?: (number | null)[] | null;
}

export interface ChartDataPointOverride {
  idx: number;
  /** Resolved fill hex (no `#`). */
  color?: string;
  markerSymbol?: string;
  markerSize?: number;
  markerFill?: string;
  markerLine?: string;
}

export interface ChartDataLabelOverride {
  idx: number;
  /** Empty string = label deleted (skip drawing). */
  text: string;
  /** "l"|"r"|"t"|"b"|"ctr"|"outEnd"|"bestFit". undefined = inherit. */
  position?: string;
  fontColor?: string;
  fontSizeHpt?: number;
  /** `<a:defRPr b="1">` inside the per-idx rich text. */
  fontBold?: boolean;
}

export interface ChartSeriesDataLabels {
  showVal: boolean;
  showCatName: boolean;
  showSerName: boolean;
  showPercent: boolean;
  position?: string;
  fontColor?: string;
  formatCode?: string;
  /** Series-level bold default for data labels. */
  fontBold?: boolean;
  /** Series-level font size for data labels (OOXML hundredths of a point). */
  fontSizeHpt?: number;
}

export interface ChartErrBars {
  /** "x" | "y". */
  dir: string;
  /** "plus" | "minus" | "both". */
  barType: string;
  plus: (number | null)[];
  minus: (number | null)[];
  noEndCap: boolean;
  /** Resolved hex (no `#`). */
  color?: string;
  lineWidthEmu?: number;
  /** "solid"|"dash"|"dot"|"dashDot"|... */
  dash?: string;
}

/**
 * Canonical chart type vocabulary. Embeds direction (`H` = horizontal) and
 * grouping (`Pct` = percent-stacked) so renderers do not need to inspect
 * separate `barDir`/`grouping` fields.
 */
export type ChartType =
  | 'line' | 'stackedLine' | 'stackedLinePct'
  | 'clusteredBar' | 'clusteredBarH'
  | 'stackedBar' | 'stackedBarH'
  | 'stackedBarPct' | 'stackedBarHPct'
  | 'area' | 'stackedArea' | 'stackedAreaPct'
  | 'pie' | 'doughnut'
  | 'scatter' | 'bubble' | 'radar' | 'waterfall'
  | string;

export interface ChartModel {
  chartType: ChartType;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
  /** Show data labels on bars / points / slices. */
  showDataLabels: boolean;
  /** Explicit Y-axis minimum (OOXML `<c:valAx><c:min>`). */
  valMin: number | null;
  /** Explicit Y-axis maximum (OOXML `<c:valAx><c:max>`). */
  valMax: number | null;
  catAxisTitle: string | null;
  valAxisTitle: string | null;
  /** `<c:catAx><c:delete val="1"/>`. */
  catAxisHidden: boolean;
  /** `<c:valAx><c:delete val="1"/>`. */
  valAxisHidden: boolean;
  /** `<c:catAx><c:spPr><a:ln><a:noFill>` — hide just the axis LINE; labels
   *  and tick marks still render. Distinct from `catAxisHidden` (which
   *  removes everything via `<c:delete val="1"/>`). */
  catAxisLineHidden: boolean;
  /** `<c:valAx><c:spPr><a:ln><a:noFill>` — hide just the axis LINE; labels
   *  and tick marks still render. */
  valAxisLineHidden: boolean;
  /** Hex without '#'. From `<c:plotArea><c:spPr><a:solidFill>`. */
  plotAreaBg: string | null;
  /** Outer chartSpace background (hex without '#'). null when noFill/absent. */
  chartBg: string | null;
  /** True when `<c:legend>` is declared in the chart XML. False = no legend. */
  showLegend: boolean;
  /** `<c:legend><c:legendPos val>` — "r"|"l"|"t"|"b"|"tr". null = default (r). */
  legendPos: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  /** `<c:catAx><c:crossBetween val="..."/>`. "between" inserts 0.5-step padding
   *  on each end of the category axis; "midCat" anchors endpoints to the axes. */
  catAxisCrossBetween: 'between' | 'midCat' | string;
  /** `<c:valAx><c:majorTickMark>`. ECMA-376 default is "cross". */
  valAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:catAx><c:majorTickMark>`. */
  catAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:valAx | catAx><c:minorTickMark>`. ECMA-376 default is "none". */
  valAxisMinorTickMark?: 'cross' | 'out' | 'in' | 'none' | string | null;
  catAxisMinorTickMark?: 'cross' | 'out' | 'in' | 'none' | string | null;
  /** Title font size in OOXML hundredths of a point (1600 = 16pt). null = default. */
  titleFontSizeHpt: number | null;
  /** Title font color as a hex string without '#' (e.g. "1B4332"). null = default. */
  titleFontColor: string | null;
  /** Title font family from `<a:latin typeface>` (ECMA-376 §20.1.4.2.24). null = default. */
  titleFontFace: string | null;
  /** `<c:catAx><c:txPr>` font size (hpt). null = fall back to proportional default. */
  catAxisFontSizeHpt: number | null;
  /** `<c:valAx><c:txPr>` font size (hpt). null = fall back to proportional default. */
  valAxisFontSizeHpt: number | null;
  /** `<c:catAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#').
   *  null = renderer default. Lets templates color category labels gray. */
  catAxisFontColor?: string | null;
  /** `<c:valAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  valAxisFontColor?: string | null;
  /** `<c:dLbls><c:txPr>` font size (hpt) for data-point value labels. */
  dataLabelFontSizeHpt: number | null;
  /** Waterfall subtotal category indices. */
  subtotalIndices: number[];
  /** `<c:legend><c:manualLayout>` absolute placement fractions of the chart
   *  space (ECMA-376 §21.2.2.31). Overrides the default side-based legend
   *  rectangle while still letting `legendPos` decide which side of the plot
   *  gets the reserved band. null = use default layout. */
  legendManualLayout?: LegendManualLayout | null;
  /**
   * `<c:valAx><c:numFmt@formatCode>` — format code applied to value-axis tick
   * labels (ECMA-376 §21.2.2.21). null = plain numeric formatting.
   */
  valAxisFormatCode?: string | null;
  /**
   * `<c:barChart><c:gapWidth>` — space between category groups as a
   * percentage of bar width (ECMA-376 §21.2.2.13). Default per spec is 150.
   * null = renderer default.
   */
  barGapWidth?: number | null;
  /**
   * `<c:barChart><c:overlap>` — signed percentage overlap between bars in the
   * same category cluster (ECMA-376 §21.2.2.25). Negative = gap, positive =
   * overlap, 0 = flush. Range [-100, 100]. null = renderer default (0).
   */
  barOverlap?: number | null;
  /**
   * `<c:dLbls><c:dLblPos>` — data label position (ECMA-376 §21.2.2.16).
   * "ctr"|"inBase"|"inEnd"|"outEnd"|"l"|"r"|"t"|"b"|"bestFit" etc.
   */
  dataLabelPosition?: string | null;
  /** Hex (no `#`) for data label text, resolved from `<c:dLbls><c:txPr>`. */
  dataLabelFontColor?: string | null;
  /**
   * `<c:dLbls><c:numFmt@formatCode>` — chart-level override for data label
   * number format (ECMA-376 §21.2.2.35). When absent, `valFormatCode` on each
   * series is used.
   */
  dataLabelFormatCode?: string | null;
  /** `<c:title>...defRPr@b>` chart title bold flag. */
  titleFontBold?: boolean | null;
  /** `<c:catAx><c:txPr>...defRPr@b>` X-axis tick label bold flag. */
  catAxisFontBold?: boolean | null;
  /** `<c:valAx><c:txPr>...defRPr@b>` Y-axis tick label bold flag. */
  valAxisFontBold?: boolean | null;
  /** `<c:catAx><c:title>` run-prop font size (hpt). Distinct from
   *  `catAxisFontSizeHpt` (tick labels). null = renderer default. */
  catAxisTitleFontSizeHpt?: number | null;
  /** `<c:catAx><c:title>` run-prop bold flag. null = not bold. */
  catAxisTitleFontBold?: boolean | null;
  /** `<c:catAx><c:title>` run-prop color (hex without '#'). null = default. */
  catAxisTitleFontColor?: string | null;
  /** `<c:valAx><c:title>` run-prop font size (hpt). null = renderer default. */
  valAxisTitleFontSizeHpt?: number | null;
  /** `<c:valAx><c:title>` run-prop bold flag. null = not bold. */
  valAxisTitleFontBold?: boolean | null;
  /** `<c:valAx><c:title>` run-prop color (hex without '#'). null = default. */
  valAxisTitleFontColor?: string | null;
  /** Explicit chart border color (hex without '#') from
   *  `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr>`. Only set when the
   *  XML explicitly declares a paintable line; null otherwise (no default
   *  border is drawn). */
  chartBorderColor?: string | null;
  /** `<c:chartSpace><c:spPr><a:ln@w>` border width in EMU. null = 1px hairline
   *  when a color is present. */
  chartBorderWidthEmu?: number | null;
  /**
   * `<c:catAx><c:crosses val>` (`autoZero` | `min` | `max`). Drives the Y
   * coordinate where the X axis is drawn. Default `autoZero` puts the X
   * axis at y=0 — that's how Excel "Project Timeline" templates split
   * milestones (positive Y) above and tasks (negative Y) below the axis.
   */
  catAxisCrosses?: string | null;
  /** `<c:catAx><c:crossesAt val>` — explicit numeric override for the
   *  crossing point. Takes precedence over `catAxisCrosses`. */
  catAxisCrossesAt?: number | null;
  valAxisCrosses?: string | null;
  valAxisCrossesAt?: number | null;
  /** Axis line color (hex without `#`) and width in EMU from
   *  `<c:catAx|valAx><c:spPr><a:ln>`. */
  catAxisLineColor?: string | null;
  catAxisLineWidthEmu?: number | null;
  valAxisLineColor?: string | null;
  valAxisLineWidthEmu?: number | null;
  /**
   * `<c:catAx><c:numFmt@formatCode>` (or scatter X-axis valAx). When set,
   * the renderer formats X-axis tick labels with this code (e.g. dates).
   */
  catAxisFormatCode?: string | null;
  /**
   * `<c:catAx><c:scaling><c:min/max>` — explicit X-axis range. Used by
   * scatter / bubble charts whose X axis is numeric. null = derive from
   * data extents.
   */
  catAxisMin?: number | null;
  catAxisMax?: number | null;
  /**
   * `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27) absolute
   * placement for the chart title.
   */
  titleManualLayout?: ChartManualLayout | null;
  /**
   * `<c:plotArea><c:layout><c:manualLayout>` absolute placement for the
   * plot area. `layoutTarget="inner"` (default) describes the inner plot
   * rect (no axes / labels); `outer` describes the outer rect (axes
   * included).
   */
  plotAreaManualLayout?: ChartManualLayout | null;
  /**
   * `<c:scatterChart><c:scatterStyle val>` (ECMA-376 §21.2.2.42). Drives
   * whether scatter charts connect points with lines and whether those
   * lines are smoothed. Values: "marker" (markers only — Excel default
   * "Scatter"), "line" / "lineMarker" (straight segments), "smooth" /
   * "smoothMarker" (cubic Bézier through points), "lineNoMarker",
   * "smoothNoMarker". null = renderer default ("marker"). Only consulted
   * for `chartType === "scatter"`; bubble ignores it.
   */
  scatterStyle?: string | null;
  /**
   * `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10). Controls
   * whether radar series render as line + markers ("standard" / "marker")
   * or as a closed polygon with area fill ("filled"). null = default
   * ("standard" — line, no fill). Only consulted for `chartType === "radar"`.
   */
  radarStyle?: string | null;
  /**
   * Secondary value axis for combo charts (bar + line). When present, series
   * with `useSecondaryAxis` are plotted against this axis's independent scale
   * and the axis is drawn on the right edge of the plot. null/absent = single
   * value axis (the common case). See {@link SecondaryValueAxis}.
   */
  secondaryValAxis?: SecondaryValueAxis | null;
  /**
   * `<c:date1904>` (ECMA-376 §21.2.2.38). When true the chart's serial
   * date-times resolve against the 1904 date system (base 1904-01-01) instead
   * of the default 1900 system. Threaded to the date formatters for date-axis
   * category labels and value-axis tick labels. Omitted/false ⇒ 1900 system.
   * Note: per §21.2.2.38 the element's `val` defaults to true when present but
   * the attribute is omitted, so `<c:date1904/>` alone means date1904=true.
   */
  date1904?: boolean;
}

/**
 * A secondary value axis (combo charts). Mirrors the primary value-axis
 * properties but lives in its own object so the flat primary-axis fields stay
 * untouched. Parsed from the right-hand `<c:valAx>` (`axPos="r"`,
 * `<c:crosses val="max">`).
 */
export interface SecondaryValueAxis {
  /** `<c:scaling><c:min val>`. null = derive from the series data. */
  min: number | null;
  /** `<c:scaling><c:max val>`. null = derive from the series data. */
  max: number | null;
  /** `<c:title>` plain text. null = no title. */
  title: string | null;
  /** `<c:delete val="1"/>` — hide labels/ticks entirely. */
  hidden: boolean;
  /** `<c:numFmt formatCode>` for tick labels. */
  formatCode?: string | null;
  /** `<c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  fontColor?: string | null;
  /** `<c:txPr>` tick-label font size (hpt). */
  fontSizeHpt?: number | null;
  /** `<c:spPr><a:ln><a:solidFill>` axis-line color (hex without '#'). */
  lineColor?: string | null;
  /** `<c:spPr><a:ln w>` axis-line width in EMU. */
  lineWidthEmu?: number | null;
  /** `<c:spPr><a:ln><a:noFill>` — hide just the axis rule. */
  lineHidden: boolean;
  /** `<c:majorTickMark>` — "cross" (default) | "out" | "in" | "none". */
  majorTickMark: string;
  /** `<c:title>` run-prop font size (hpt). */
  titleFontSizeHpt?: number | null;
  /** `<c:title>` run-prop bold flag. */
  titleFontBold?: boolean | null;
  /** `<c:title>` run-prop color (hex without '#'). */
  titleFontColor?: string | null;
}

/**
 * `<c:manualLayout>` block. Fractions are of the chart-space rect.
 * `xMode`/`yMode`: "edge" = absolute fraction from top-left, "factor" =
 * fraction offset from default position.
 */
export interface ChartManualLayout {
  xMode: string;
  yMode: string;
  layoutTarget?: string;
  x: number;
  y: number;
  w?: number;
  h?: number;
}

export interface LegendManualLayout {
  /** `"edge"` = `x`/`y` are fractions from top-left of chart space;
   *  `"factor"` = fractions offset from the default position. */
  xMode: string;
  yMode: string;
  /** Fractions of chart space width/height. */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ChartRect {
  x: number;
  y: number;
  w: number;
  h: number;
}
