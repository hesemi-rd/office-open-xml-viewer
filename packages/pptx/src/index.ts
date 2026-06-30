export { PptxViewer, type PptxViewerOptions, type HiddenSlideMode } from './viewer';
export {
  PptxPresentation,
  type LoadOptions,
  type RenderSlideOptions,
  type RenderSlideToBitmapOptions,
} from './presentation';
export { renderSlide, type RenderOptions, type PptxTextRunInfo, type TextRunCallback } from './renderer';
export type { PresentationHandle } from './presentation-handle';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
export type {
  Presentation,
  Slide,
  // Reachable via Slide.comments — exported so consumers reading legacy slide
  // comments have a name for the element type.
  PptxComment,
  SlideElement,
  ShapeElement,
  PictureElement,
  // SlideElement union members — reachable via Slide.elements; exported so a
  // consumer that narrows on `el.type` has a name for every variant.
  TableElement,
  TableRow,
  TableCell,
  ChartElement,
  MediaElement,
  TextRect,
  // 3D scene types (reachable via ShapeElement.scene3d / .sp3d).
  Scene3d,
  Camera3d,
  Rot3d,
  LightRig,
  Sp3d,
  Bevel3d,
  // Fill / stroke variants (reachable via ShapeElement.fill / .stroke etc.).
  Fill,
  SolidFill,
  NoFill,
  GradientFill,
  GradientStop,
  ImageFill,
  FillRect,
  TileInfo,
  Stroke,
  // Effect types (reachable via ShapeElement.shadow / .glow / …).
  Shadow,
  Glow,
  SoftEdge,
  Reflection,
  // Geometry + text run sub-types (reachable via ShapeElement / TextBody).
  PathCmd,
  Bullet,
  // Picture-bullet variant (`<a:buBlip>`, §21.1.2.4.2) — part of the PPTX
  // `Bullet` union; exported so consumers narrowing on `bullet.type === 'blip'`
  // have a name for the shape.
  BlipBullet,
  SpaceLine,
  TabStop,
  TextBody,
  Paragraph,
  TextRun,
  TextRunData,
  LineBreak,
  // Chart model (reachable via ChartElement.series and renderChart).
  ChartModel,
  ChartSeries,
  // Hidden-slide dimming options (reachable via RenderSlideOptions.dim /
  // RenderSlideToBitmapOptions.dim) — the translucent overlay mechanism.
  DimOptions,
} from './types';
