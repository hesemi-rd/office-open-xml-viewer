export type {
  ArrowEnd,
  Bullet,
  Fill,
  FillRect,
  Glow,
  GradientFill,
  GradientStop,
  ImageFill,
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
  TileInfo,
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
  buildPresetGeometryPath,
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
// DrawingML 3D camera perspective projection (planar homography).
// ECMA-376 §20.1.5.5 (camera) / §20.1.5.11 (rot). sp3d bevel / extrusion / light
// shading lives in ./shape/bevel-shading (exported below).
export {
  computeScene3dQuad,
  isScene3dNonIdentity,
  computeDepthOffset,
  type CameraInput,
  type RotInput,
  type Scene3dQuad,
  type Vec2,
} from './shape/scene3d-camera';
export { drawProjected, expandProjectedQuad } from './shape/scene3d-draw';
// DrawingML 3D bevel shading (Phase B). ECMA-376 §20.1.5.12 (sp3d) /
// §20.1.5.3 (bevel) / §20.1.10.9 (ST_BevelPresetType) / §20.1.5.9 (lightRig).
export {
  applyBevelShading,
  applyExtrusion,
  type ExtrusionInput,
  computeBevelNormals,
  bevelHeightProfile,
  distanceToEdge,
  edt1d,
  shadePixel,
  shadeParamsFor,
  lightDirFromRig,
  materialClass,
  type BevelMaterial,
  type BevelShadeParams,
  type BevelInput,
  type BevelCtx,
  type Vec3,
} from './shape/bevel-shading';
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
