export type {
  ArrowEnd,
  Bullet,
  Fill,
  Glow,
  GradientFill,
  GradientStop,
  LineBreak,
  NoFill,
  Paragraph,
  PathCmd,
  PatternFill,
  Reflection,
  RenderOptions,
  Shadow,
  SoftEdge,
  SolidFill,
  SpaceLine,
  Stroke,
  TabStop,
  TextBody,
  TextOutline,
  TextRun,
  TextRunData,
} from './types/common';
export type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartErrBars,
  ChartManualLayout,
  ChartModel,
  ChartRect,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartType,
  LegendManualLayout,
} from './types/chart';
export type { LoadOptions } from './types/load-options';
export { preloadGoogleFonts, type FontPreloadEntry } from './fonts/preload';
export { renderChart } from './chart/renderer';
export { autoResize, type AutoResizeOptions } from './autoResize';
export { buildCustomPath } from './shape/custGeom';
export { hexToRgba, resolveFill, applyStroke } from './shape/paint';
export { buildShapePath, drawStar, drawPolygon, ooxmlArcTo } from './shape/preset';
// ECMA-376 §20.1.9 spec-driven preset geometry engine (presets.json from
// presetShapeDefinitions.xml). Coexists with the legacy hand-rolled
// `buildShapePath` above, which the pptx renderer still uses as a silhouette /
// fallback codepath. Consolidating the two is intentionally out of scope here.
export {
  renderPresetShape,
  hasPreset,
  getConnectorAnchors,
} from './shape/preset-geometry';
export { type PresetPath } from './shape/preset-geometry/path-executor';
export {
  applyInnerShadow,
  applySoftEdge,
  applyReflection,
  createAuxCanvas,
  type PaintShape,
  type EffectBBox,
} from './shape/effects';
export {
  renderSparkline,
  type SparklineKind,
  type SparklineModel,
  type SparklineRect,
} from './sparkline/renderer';
export {
  mathToMathML,
  svgExtents,
  recolorSvg,
  type MathSvg,
  type MathRenderer,
} from './math';
export type { MathNode, MathFormula, MathStyle } from './types/math';
export { EMU_PER_INCH, EMU_PER_PT, EMU_PER_PX, PT_TO_PX } from './units';
export {
  WorkerBridge,
  type WorkerLike,
  type WorkerBridgeOptions,
  decodeDataUrl,
} from './worker';
export {
  toVisualSegments,
  resolveBaseDirection,
  getDefaultBidiEngine,
  setBidiEngine,
  resetBidiEngine,
  type BidiEngine,
  type BaseDirection,
  type BidiClass,
  type BidiLevels,
  type StyledRun,
  type VisualSegment,
  type SegmentPart,
} from './text/bidi';
