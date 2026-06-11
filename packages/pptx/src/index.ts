export { PptxViewer, type PptxViewerOptions } from './viewer';
export { PptxPresentation, type LoadOptions, type RenderSlideOptions } from './presentation';
export { renderSlide, type RenderOptions, type PptxTextRunInfo, type TextRunCallback } from './renderer';
export type { PresentationHandle } from './presentation-handle';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
export type {
  Presentation,
  Slide,
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
  // Fill / stroke variants (reachable via ShapeElement.fill / .stroke etc.).
  Fill,
  SolidFill,
  NoFill,
  GradientFill,
  GradientStop,
  Stroke,
  // Effect types (reachable via ShapeElement.shadow / .glow / …).
  Shadow,
  Glow,
  SoftEdge,
  Reflection,
  // Geometry + text run sub-types (reachable via ShapeElement / TextBody).
  PathCmd,
  Bullet,
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
} from './types';
