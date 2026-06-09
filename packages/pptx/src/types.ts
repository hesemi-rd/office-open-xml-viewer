// ===== Shared types re-exported from @silurus/ooxml-core =====
export type {
  PathCmd,
  Fill, SolidFill, NoFill, GradientFill, GradientStop,
  Shadow,
  Glow,
  SoftEdge,
  Reflection,
  Stroke,
  TextBody,
  SpaceLine,
  Bullet,
  TabStop,
  Paragraph,
  TextRun, TextRunData, LineBreak,
  RenderOptions,
  ChartModel, ChartSeries,
} from '@silurus/ooxml-core';

// ===== Presentation data model =====
// All positions and sizes are in EMUs (English Metric Units).
// 914400 EMU = 1 inch, 12700 EMU = 1 pt

import type { Fill, Stroke, TextBody, Shadow, Glow, SoftEdge, Reflection, PathCmd, ChartSeries } from '@silurus/ooxml-core';

export interface Presentation {
  slideWidth: number;
  slideHeight: number;
  slides: Slide[];
  /** Theme dk1 color (e.g. "383838"). Used as fallback text color when no explicit color is set. */
  defaultTextColor: string | null;
  /** Theme major (heading) font family name (e.g. "Aptos Display", "Nunito Sans"). Null if not set. */
  majorFont: string | null;
  /** Theme minor (body) font family name (e.g. "Aptos", "Nunito Sans"). Null if not set. */
  minorFont: string | null;
  /** Theme hyperlink colour (hex 6 chars). Used to colour hyperlink runs that have no explicit colour. */
  hlinkColor?: string;
  /** Theme followed-hyperlink colour. Reserved for future visited-link styling. */
  folHlinkColor?: string;
}

export interface Slide {
  index: number;
  /** 1-based slide number (index + 1); used to render slidenum fields */
  slideNumber: number;
  background: Fill | null;
  elements: SlideElement[];
}

export type SlideElement = ShapeElement | PictureElement | TableElement | ChartElement | MediaElement;

export interface MediaElement {
  type: 'media';
  x: number;
  y: number;
  width: number;
  height: number;
  /** "audio" or "video" */
  mediaKind: 'audio' | 'video';
  /** Poster image zip path (e.g. "ppt/media/image2.png"). Empty when no poster. */
  posterPath: string;
  /** Poster image MIME type (empty when no poster). */
  posterMimeType: string;
  /** Path inside the pptx zip (e.g. "ppt/media/media2.mp4"). Used by getMedia. */
  mediaPath: string;
  /** MIME type of the underlying media (e.g. "audio/mpeg", "video/mp4"). */
  mimeType: string;
}

export interface ShapeElement {
  type: 'shape';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Rotation in degrees, clockwise */
  rotation: number;
  /** Horizontal mirror (a:xfrm flipH) */
  flipH: boolean;
  /** Vertical mirror (a:xfrm flipV) */
  flipV: boolean;
  /** OOXML preset name or "custGeom" when custom paths are used */
  geometry: string;
  fill: Fill | null;
  stroke: Stroke | null;
  textBody: TextBody | null;
  /** Default text color from p:style > fontRef (hex). Used when run/para has no explicit color. */
  defaultTextColor: string | null;
  /** Custom geometry sub-paths (set only when geometry === "custGeom").
   *  Outer array: one entry per <a:path>; inner: path commands with coords in [0,1]. */
  custGeom: PathCmd[][] | null;
  /** First adjustment value from prstGeom avLst (e.g. trapezoid inset). Range 0–100000. */
  adj: number | null;
  /** Second adjustment value from prstGeom avLst (e.g. arrow head width). Range 0–100000. */
  adj2: number | null;
  /** Third adjustment value from prstGeom avLst (e.g. callout tip x). Range 0–100000. */
  adj3: number | null;
  /** Fourth adjustment value from prstGeom avLst (e.g. callout tip y). Range 0–100000. */
  adj4: number | null;
  /** adj5-adj8: extra polyline vertices for callouts like accentBorderCallout3. */
  adj5: number | null;
  adj6: number | null;
  adj7: number | null;
  adj8: number | null;
  /** Drop shadow from effectLst > outerShdw (null if not present). */
  shadow: Shadow | null;
  /** Inner (inset) shadow from effectLst > innerShdw. ECMA-376 §20.1.8.21. */
  innerShadow?: Shadow;
  /** Coloured glow halo from effectLst > glow. ECMA-376 §20.1.8.17. */
  glow?: Glow;
  /** Soft (feathered) edge — ECMA-376 §20.1.8.31. */
  softEdge?: SoftEdge;
  /** Mirrored reflection — ECMA-376 §20.1.8.27. */
  reflection?: Reflection;
  /** Explicit text frame from a SmartArt drawing's `<dsp:txXfrm>` (absolute EMU,
   *  same space as x/y/width/height). When present the renderer lays text out in
   *  this rectangle instead of the preset/ellipse-derived text rectangle. */
  textRect?: TextRect;
}

/** Absolute text-frame rectangle in EMU (from SmartArt `<dsp:txXfrm>`). */
export interface TextRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TableElement {
  type: 'table';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Column widths in EMU */
  cols: number[];
  rows: TableRow[];
}

export interface TableRow {
  /** Row height in EMU */
  height: number;
  cells: TableCell[];
}

export interface TableCell {
  textBody: TextBody | null;
  fill: Fill | null;
  /** Default run text colour inherited from the table style (`<a:tcTxStyle>`); hex, no `#`. */
  textColor?: string;
  borderL: Stroke | null;
  borderR: Stroke | null;
  borderT: Stroke | null;
  borderB: Stroke | null;
  /** Diagonal from top-left to bottom-right */
  diagonalTL?: Stroke | null;
  /** Diagonal from top-right to bottom-left */
  diagonalTR?: Stroke | null;
  /** Column span */
  gridSpan: number;
  /** Row span */
  rowSpan: number;
  /** Horizontal merge continuation */
  hMerge: boolean;
  /** Vertical merge continuation */
  vMerge: boolean;
}

/**
 * PPTX chart element. The Rust parser emits ChartModel fields flat at the
 * top level, alongside the element position (x/y/width/height in EMU).
 * Pass this straight to `renderChart` from `@silurus/ooxml-core`.
 */
export interface ChartElement {
  type: 'chart';
  x: number;
  y: number;
  width: number;
  height: number;
  chartType: string;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
  valMax: number | null;
  valMin: number | null;
  subtotalIndices: number[];
  showDataLabels: boolean;
  catAxisHidden: boolean;
  valAxisHidden: boolean;
  /** `<c:catAx><c:spPr><a:ln><a:noFill>` — line-only hide; labels stay. */
  catAxisLineHidden?: boolean;
  /** `<c:valAx><c:spPr><a:ln><a:noFill>` — line-only hide; labels stay. */
  valAxisLineHidden?: boolean;
  plotAreaBg: string | null;
  /** Outer chartSpace background (hex without '#'). null when noFill/absent. */
  chartBg: string | null;
  /** True when <c:legend> is declared; false suppresses the legend entirely. */
  showLegend: boolean;
  /** catAx crossBetween: "between" (default, 0.5-step padding) or "midCat". */
  catAxisCrossBetween: 'between' | 'midCat' | string;
  /** `<c:valAx><c:majorTickMark>`. "cross" (default) | "out" | "in" | "none". */
  valAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:catAx><c:majorTickMark>`. */
  catAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** Title font size in OOXML hundredths of a point (1600 = 16pt). null = default. */
  titleFontSizeHpt: number | null;
  /** Title font color as a hex string without '#'. null = default/theme. */
  titleFontColor?: string | null;
  /** Title font family (`<a:latin typeface>`). null = default/theme. */
  titleFontFace?: string | null;
  /** `<c:catAx><c:txPr>` font size (hpt). null = proportional default. */
  catAxisFontSizeHpt: number | null;
  /** `<c:valAx><c:txPr>` font size (hpt). null = proportional default. */
  valAxisFontSizeHpt: number | null;
  /** `<c:catAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  catAxisFontColor?: string | null;
  /** `<c:valAx><c:txPr>…<a:solidFill>` tick-label color (hex without '#'). */
  valAxisFontColor?: string | null;
  /** `<c:catAx><c:spPr><a:ln><a:solidFill>` axis-line color (hex without '#'). */
  catAxisLineColor?: string | null;
  /** `<c:catAx><c:spPr><a:ln w>` axis-line width in EMU. */
  catAxisLineWidthEmu?: number | null;
  /** `<c:valAx><c:spPr><a:ln><a:solidFill>` axis-line color (hex without '#'). */
  valAxisLineColor?: string | null;
  /** `<c:valAx><c:spPr><a:ln w>` axis-line width in EMU. */
  valAxisLineWidthEmu?: number | null;
  /** `<c:dLbls><c:txPr>` font size (hpt) for data-point value labels. */
  dataLabelFontSizeHpt: number | null;
  /** `<c:legend><c:legendPos val>` — "r" (default) | "l" | "t" | "b" | "tr". */
  legendPos?: 'r' | 'l' | 't' | 'b' | 'tr' | null;
  /** `<c:barChart><c:gapWidth val>` — % of bar width between category groups (default 150). */
  barGapWidth?: number | null;
  /** `<c:barChart><c:overlap val>` — signed % of bar width for cluster overlap. */
  barOverlap?: number | null;
  /** `<c:dLbls><c:dLblPos val>` — data label placement ("ctr" | "inEnd" | "outEnd" | …). */
  dataLabelPosition?: string | null;
  /** `<c:dLbls><c:txPr>…<a:solidFill>` resolved to hex (no '#'). null = renderer default. */
  dataLabelFontColor?: string | null;
  /** `<c:dLbls><c:numFmt formatCode>` — data label number format. */
  dataLabelFormatCode?: string | null;
  /** `<c:valAx><c:numFmt formatCode>` — value-axis tick label number format. */
  valAxisFormatCode?: string | null;
  /** `<c:plotArea><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.32) — explicit
   * plot-area placement so bars don't extend past the chart-frame's intended
   * inner region (sample-2 slide-16 horizontal bar chart). */
  plotAreaManualLayout?: import('@silurus/ooxml-core').ChartManualLayout | null;
  /** `<c:scatterChart><c:scatterStyle val>` (ECMA-376 §21.2.2.42) — drives
   * whether scatter charts connect points with straight or smooth lines. */
  scatterStyle?: string | null;
  /** `<c:radarChart><c:radarStyle val>` (ECMA-376 §21.2.3.10). */
  radarStyle?: string | null;
}

export interface PictureElement {
  type: 'picture';
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /** Data URL, e.g. "data:image/png;base64,..." */
  dataUrl: string;
  /** OOXML adj value (0–100000) for roundRect clip, null = plain rectangle */
  clipAdjust: number | null;
  /**
   * ECMA-376 a:srcRect — source image crop as fractions (0..1) of the source
   * width/height. Omitted when the image is not cropped.
   */
  srcRect?: { l?: number; t?: number; r?: number; b?: number };
  /** a:blip > a:alphaModFix@amt as 0..1. Undefined = fully opaque. */
  alpha?: number;
  /**
   * `<p:spPr><a:custGeom>` clipping path. Same `PathCmd` model as
   * `ShapeElement.custGeom` (one entry per `<a:path>`; coords normalized
   * into [0,1] of the picture's bounding box). The renderer builds a
   * Path2D and `ctx.clip()` before drawing the bitmap so the image is
   * trimmed to the laptop / device silhouette declared in the file.
   */
  custGeom?: PathCmd[][] | null;
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { kind: 'init'; wasmUrl: string }
  | { kind: 'parse'; id: number; buffer: ArrayBuffer; maxZipEntryBytes?: number }
  | { kind: 'extractMedia'; id: number; path: string };

export type WorkerResponse =
  | { kind: 'ready' }
  | { kind: 'parsed'; id: number; presentation: Presentation }
  | { kind: 'mediaExtracted'; id: number; bytes: ArrayBuffer }
  | { kind: 'error'; id: number; message: string };
