export { XlsxWorkbook, type LoadOptions } from './workbook.js';
export { XlsxViewer } from './viewer.js';
export type { XlsxViewerOptions, CellAddress, CellRange, SelectionMode } from './viewer.js';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
export type {
  Workbook,
  SheetMeta,
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
  // Embedded chart model sub-types.
  ChartData,
  XlsxChartSeries,
  SeriesDataLabels,
  DataLabelOverride,
  DataPointOverride,
  ErrBars,
  ManualLayout,
  LegendManualLayout,
} from './types.js';
