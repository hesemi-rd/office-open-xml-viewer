export { XlsxWorkbook, type LoadOptions } from './workbook.js';
export type { WireRenderViewportOptions } from './worker-protocol.js';
export { XlsxViewer } from './viewer.js';
// Resolved list-validation values (reachable via XlsxWorkbook.resolveValidationList).
export type { ResolvedList } from './validation-list.js';
export type { XlsxViewerOptions, HiddenSheetMode, CellAddress, CellRange, SelectionMode } from './viewer.js';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
// Resolve `{type:'shared',si}` cells against a workbook's sharedStrings table
// (ECMA-376 §18.4.8). Exported so headless callers that parse a Worksheet
// directly (e.g. @silurus/ooxml-node's parseXlsxSheet) can concretize cell text.
export { resolveSharedStrings } from './shared-strings.js';
// Typed load-time error surfaced by XlsxWorkbook.load (e.g. a password-protected
// or legacy-binary .xls file). Re-exported so `@silurus/ooxml/xlsx` consumers can
// narrow on `err.code`.
export { OoxmlError, type OoxmlErrorCode } from '@silurus/ooxml-core';
export type {
  Workbook,
  SheetMeta,
  SheetVisibility,
  Worksheet,
  Row,
  Cell,
  CellValue,
  Styles,
  CellFont,
  CellFill,
  Border,
  BorderEdge,
  CellXf,
  NumFmt,
  MergeCell,
  ParsedWorkbook,
  ViewportRange,
  RenderViewportOptions,
  XlsxTextRunInfo,
  // Rich-text run sub-types (reachable via Cell rich-text values).
  Run,
  RunFont,
  SharedString,
  // Differential / gradient style sub-types (reachable via Styles).
  Dxf,
  GradientFillSpec,
  // Conditional formatting (reachable via Worksheet.conditionalFormats).
  ConditionalFormat,
  CfRule,
  CfValue,
  CfStop,
  CfIcon,
  // Workbook-level metadata.
  DefinedName,
  Hyperlink,
  // Cell comments / notes (reachable via Worksheet.comments).
  XlsxComment,
  // Data validation rules (reachable via Worksheet.dataValidations).
  DataValidation,
  // Excel tables (reachable via Worksheet.tables).
  TableInfo,
  TableColumnInfo,
  // Slicers.
  SlicerAnchor,
  SlicerItem,
  // Sparklines (reachable via Worksheet sparkline groups).
  SparklineGroup,
  Sparkline,
  // Drawings / shapes (reachable via Worksheet drawings).
  ImageAnchor,
  ChartAnchor,
  ShapeAnchor,
  ShapeInfo,
  ShapeGeom,
  ShapeText,
  ShapeParagraph,
  ShapeTextRun,
  PathInfo,
  PathCmd,
  // Canonical chart model (shared with core / pptx). `ChartAnchor.chart` is a
  // `ChartModel`.
  ChartModel,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartErrBars,
  ChartManualLayout,
  LegendManualLayout,
  // Back-compat aliases for the former XLSX-local chart types (now the core
  // sub-types). `ChartData` is removed — it described the pre-adapter parse
  // shape, which is no longer emitted.
  XlsxChartSeries,
  SeriesDataLabels,
  DataLabelOverride,
  DataPointOverride,
  ErrBars,
  ManualLayout,
} from './types.js';
