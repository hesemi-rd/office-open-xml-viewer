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
  SecondaryValueAxis,
} from './types/chart';
export type { LoadOptions } from './types/load-options';
export { preloadGoogleFonts, type FontPreloadEntry } from './fonts/preload';
export {
  classifyCjkFont,
  classifyFontGeneric,
  isComplexScriptCodePoint,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  SCRIPT_GOOGLE_FONTS,
  SCRIPT_PRELOAD_NAMES,
  scriptPreloadNamesForText,
  type CjkLang,
  type FontGenericClass,
  type FontVariant,
} from './fonts/scripts';
export {
  SYMBOL_MAP,
  WINGDINGS_MAP,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  type SymbolTextSegment,
} from './fonts/symbol-font';
export { renderChart } from './chart/renderer';
export { autoResize, type AutoResizeOptions } from './autoResize';
export { buildCustomPath } from './shape/custGeom';
export {
  getCustGeomEndpoints,
  type CustGeomEndpoint,
  type CustGeomEndpoints,
} from './shape/custgeom-endpoints';
export { hexToRgba, relativeLuma, autoContrastColor, resolveFill, applyStroke } from './shape/paint';
export { buildShapePath, drawStar, drawPolygon, ooxmlArcTo } from './shape/preset';
export { drawArrowHead } from './shape/arrow';
// Shared embedded-SVG decoder (Microsoft asvg:svgBlip extension) — used by all
// three renderers to prefer the vector original over the raster fallback.
// Path-keyed for the lazy byte-on-demand pipeline: fetches SVG bytes via a
// caller-supplied fetchImage(path, mimeType) and owns the object-URL lifecycle.
export { getCachedSvgImageByPath, dropSvgImageCache } from './image/svg-image-by-path';
// Shared WMF (Windows Metafile) player + the raster/metafile decoder all three
// renderers route through, so a WMF/EMF blip (which `createImageBitmap` cannot
// decode) rasterizes or is skipped gracefully instead of throwing and vanishing.
// The docx-specific cosmetic window/device-frame suppression is gated behind
// `suppressBoundaryFrame` (default off = spec-clean).
export {
  isWmf,
  isEmf,
  isMetafileMime,
  playWmf,
  renderWmfToBitmap,
  wmfRasterTarget,
  decodeRasterOrMetafile,
  type DecodeRasterOptions,
} from './image/wmf';
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
  fillDirFromKey,
  materialClass,
  type BevelMaterial,
  type BevelShadeParams,
  type LightRigRot,
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
export { isHTMLCanvas, defaultDpr } from './canvas/env';
export { crispOffset } from './canvas/crisp';
// Shared border / line dash-pattern core (§17.18.2 ST_Border / §18.18.3
// ST_BorderStyle / §20.1.10.49 ST_PresetLineDashVal). One logical vocabulary +
// the [on, off, …].map(x => x*unit) helper + per-format relative tables; each
// format keeps its own multipliers (output is byte-identical to the old inline
// implementations).
export {
  dashArray,
  docxBorderDashArray,
  xlsxBorderDashArray,
  pptxDashArray,
  type RelativeDashPattern,
} from './draw/dash';
// Shared `double` border rail geometry (§17.18.2 / §18.18.3): floored-thirds
// device-pixel rail/gap/rail bands + a fill-based painter.
export { doubleRailGeometry, fillDoubleBorder } from './draw/double-border';
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
export type { KinsokuRules } from './text/kinsoku';
export {
  resolveKinsokuRules,
  DEFAULT_KINSOKU_RULES,
  kinsokuAdjustedSplit,
  crossRunKinsokuRetract,
} from './text/kinsoku';
export { isCjkBreakChar } from './text/cjk-ranges';
export { highlightBox } from './text/highlight-box';
export {
  distributeLineSlack,
  type DistributeSeg,
  type DistributeResult,
  type DistributeOptions,
  type SegStretch,
} from './text/line-distribute';
export { justifiedPiecePositions, type JustifiedPiece } from './text/justify-positions';
