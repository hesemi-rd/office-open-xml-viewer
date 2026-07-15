import type { SectionLayoutContext } from '../layout-context.js';
import type {
  TextFontSlotPresence,
  TextFontSlots,
  TextLayoutService,
} from './text.js';
import type { ImageMetadataService, MathMetadataService } from './resources.js';
import type { AnchorFrameResult } from './anchor-frame.js';
import type {
  CanvasFontRoute,
  ChartModel,
  DrawingMLShapePaintPlan,
  Duotone,
  Fill,
} from '@silurus/ooxml-core';

export type { TextLayoutService } from './text.js';
export type { ImageMetadataService, MathMetadataService } from './resources.js';

export type LayoutNodeId = string;

export type SourceRef = Readonly<{
  story: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textbox';
  storyInstance: string;
  path: readonly number[];
}>;

export type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
  : T extends readonly (infer U)[] ? readonly DeepReadonly<U>[]
  : T extends object ? { readonly [K in keyof T]: DeepReadonly<T[K]> }
  : T;

export interface PointPt {
  readonly xPt: number;
  readonly yPt: number;
}

export interface LayoutRect extends PointPt {
  readonly widthPt: number;
  readonly heightPt: number;
}

export type FlowDomainKind =
  | 'body'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'textbox'
  | 'tableCell';

export interface FlowDomain {
  readonly id: string;
  readonly kind: FlowDomainKind;
  readonly bounds: LayoutRect;
}

export interface PageGeometry extends LayoutRect {
  readonly contentTopPt: number;
  readonly contentBottomPt: number;
}

export interface Matrix2DData {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export type ClipPathData =
  | Readonly<{ kind: 'rect'; rect: LayoutRect }>
  | Readonly<{ kind: 'polygon'; points: readonly PointPt[] }>;

export interface FlowOwnership {
  readonly flowDomainId: string;
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly advancePt: number;
  readonly ordinaryFlow: boolean;
}

interface LayoutNodeBase extends FlowOwnership {
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
}

export type DrawingPaintCommand =
  | Readonly<{
      kind: 'noop';
    }>
  | Readonly<{
      kind: 'drawingml-shape';
      plan: DeepReadonly<DrawingMLShapePaintPlan>;
    }>
  | Readonly<{
      kind: 'fill-rect';
      rect: LayoutRect;
      fill: string;
    }>
  | Readonly<{
      kind: 'stroke-rect';
      rect: LayoutRect;
      stroke: string;
      lineWidthPt: number;
      dashPt: readonly number[];
    }>
  | Readonly<{
      kind: 'text';
      rect: LayoutRect;
      text: string;
      fill: string;
      fontRoute: CanvasFontRoute;
      fontSizePt: number;
      fontWeight: number;
      fontStyle: 'normal' | 'italic';
      align: 'start' | 'center' | 'end';
      baseline: 'top' | 'middle' | 'alphabetic' | 'bottom';
    }>
  | Readonly<{
      kind: 'watermark-text';
      rect: LayoutRect;
      text: string;
      fill: DeepReadonly<Fill> | null;
      opacity: number;
      rotationDeg: number;
      /** True applies §19.1.2.23 fitshape; false preserves authored font size. */
      fitShape: boolean;
      fontSizePt: number;
      /** Glyph source box relative to span origin x=0 / alphabetic baseline y=0. */
      sourceBounds: LayoutRect;
      spans: readonly Readonly<{
        text: string;
        advancePt: number;
        fontRoute: CanvasFontRoute;
        fontWeight: number;
        fontStyle: 'normal' | 'italic';
      }>[];
    }>
  | Readonly<{
      kind: 'resource';
      resourceKey: string;
      resourceKind: PaintResourceKind;
      rect: LayoutRect;
    }>;

export interface DrawingLayout extends LayoutNodeBase {
  readonly kind: 'drawing';
  readonly transform?: Matrix2DData;
  readonly clip?: ClipPathData;
  readonly commands: readonly DrawingPaintCommand[];
  readonly anchorLayer?: Readonly<{
    occurrenceId: string;
    behindDoc: boolean;
    relativeHeight: number;
    sourceOrder: number;
    horizontalOwnership: 'page' | 'host';
    verticalOwnership: 'page' | 'host';
  }>;
  readonly textBoxIds?: readonly LayoutNodeId[];
}

/** Clone-safe transitional VML facts projected at the parser/model boundary.
 * Presence of the boolean controls distinguishes parser-created false defaults
 * from the stable public ShapeRun compatibility surface. */
export interface VmlTextPathAcquisitionInput {
  readonly string: string;
  readonly fontFamily?: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly textPathOk?: boolean;
  readonly on?: boolean;
  readonly fitShape?: boolean;
  readonly fitPath?: boolean;
  readonly trim?: boolean;
  readonly xScale?: boolean;
  readonly fontSizePt?: number;
}

export interface TextRange {
  readonly start: number;
  readonly end: number;
}

export type TextDirection = 'ltr' | 'rtl';
export type WritingMode = 'horizontal-tb' | 'vertical-rl' | 'vertical-lr';

export type TextDecorationLayout = Readonly<{
  kind: 'underline' | 'strikethrough' | 'overline';
  /** Original ECMA-376 ST_Underline token when this is a w:u operation. */
  authoredStyle?: string;
  from: PointPt;
  to: PointPt;
  color: string;
  widthPt: number;
  style: 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy';
  /** Final acquired path. Multi-stroke/dash/wave expansion belongs to layout. */
  path?: readonly PointPt[];
  readonly dashPatternPt?: readonly number[];
}>;

export interface RetainedGlyphPaintOperation {
  readonly text: string;
  readonly origin: PointPt;
  readonly fontRoute: CanvasFontRoute;
  readonly fontSizePt: number;
  readonly fontWeight: number;
  readonly fontStyle: 'normal' | 'italic';
  readonly color: TextColorPolicy;
  /** Tight selected-face ink relative to this operation's baseline origin. */
  readonly inkBounds?: Readonly<{
    xMinPt: number;
    xMaxPt: number;
    ascentPt: number;
    descentPt: number;
  }>;
}

export type RetainedMarkPath = Readonly<{
  kind: 'polyline';
  points: readonly PointPt[];
  fill: string | null;
  stroke: string | null;
  strokeWidthPt: number;
}>;

export interface RetainedRunBorderFacts {
  readonly val: string;
  readonly color: string;
  readonly widthPt: number;
  readonly spacePt: number;
  readonly themeColor?: string;
  readonly themeTint?: string;
  readonly themeShade?: string;
  readonly shadow?: boolean;
  readonly frame?: boolean;
}

export interface TextClusterLayout {
  readonly range: TextRange;
  readonly offset: PointPt;
  readonly advancePt: number;
}

export interface TextPaintOp {
  readonly text: string;
  readonly range: TextRange;
  readonly offset: PointPt;
  readonly letterSpacingPt: number;
  readonly scaleX: number;
  readonly direction: TextDirection;
  readonly kerning: 'auto' | 'normal' | 'none';
  readonly writingMode: WritingMode;
  readonly glyphOrientation?: 'sideways' | 'upright';
  /** `kashida` permits acquisition-inserted U+0640 glyphs over one source range. */
  readonly sourceMapping?: 'identity' | 'kashida';
}

export type TextColorPolicy =
  | Readonly<{ kind: 'explicit'; color: string }>
  | Readonly<{ kind: 'auto'; background?: string }>
  | Readonly<{ kind: 'default' }>;

export type RetainedTypographyValue<T> = Readonly<{
  status: 'missing' | 'invalid' | 'valid';
  raw: string | null;
  value: T | null;
}>;

export interface RetainedRunTypographyFacts {
  readonly caps: boolean;
  readonly smallCaps: boolean;
  readonly strike: boolean;
  readonly doubleStrike: boolean;
  readonly verticalAlign: RetainedTypographyValue<'super' | 'sub'>;
  readonly positionPt: RetainedTypographyValue<number>;
  readonly emphasis: RetainedTypographyValue<string>;
  readonly underline?: Readonly<{
    val: RetainedTypographyValue<string>;
    color: RetainedTypographyValue<string>;
    themeColor: RetainedTypographyValue<string>;
    themeTint: RetainedTypographyValue<string>;
    themeShade: RetainedTypographyValue<string>;
  }>;
}

export interface TextPlacement {
  readonly kind: 'text';
  readonly text: string;
  /** Parsed run occurrence retained for destination-page field convergence. */
  readonly sourceRunIndex?: number;
  readonly role?: 'content' | 'numbering-marker' | 'field-result';
  readonly dependency?: 'page' | 'total-pages' | 'date' | 'time' | 'document';
  readonly range: TextRange;
  readonly origin: PointPt;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
  /** Shaped cluster geometry for selection/hit testing. Always covers `range`. */
  readonly clusters: readonly TextClusterLayout[];
  /** Immutable contextual paint operations. Normally one whole-run operation. */
  readonly paintOps: readonly TextPaintOp[];
  readonly color: TextColorPolicy;
  readonly fontRoute: CanvasFontRoute;
  readonly fontSizePt: number;
  readonly fontWeight: number;
  readonly fontStyle: 'normal' | 'italic';
  readonly direction: TextDirection;
  readonly writingMode?: WritingMode;
  readonly characterSpacingPt?: number;
  readonly characterScale?: number;
  readonly fitText?: Readonly<{ regionIndex: number; perGapPt: number; trailingPadPt: number }>;
  readonly kerning?: boolean;
  readonly positionPt?: number;
  readonly verticalAlign?: 'super' | 'sub';
  readonly tateChuYoko?: boolean;
  readonly tateChuYokoCompress?: boolean;
  readonly ruby?: Readonly<{
    text: string;
    advancePt: number;
    authored: Readonly<{
      align?: string;
      baseFontSizePt?: number;
      raisePt?: number;
      language?: string;
    }>;
    readonly paintOps: readonly RetainedGlyphPaintOperation[];
  }>;
  readonly emphasisMark?: string;
  readonly emphasis?: Readonly<{
    authored: string;
    /** Selected authored mark glyphs, one per non-space source cluster. */
    glyphs?: readonly RetainedGlyphPaintOperation[];
    /** Authoritative outline paths remain representable when supplied by a font service. */
    paths?: readonly RetainedMarkPath[];
  }>;
  readonly highlight?: string;
  readonly highlightFragments?: readonly Readonly<{ rect: LayoutRect; color: string }>[];
  readonly background?: string;
  /** Justification width owned after this visual fragment. */
  readonly ownedTrailingSlackPt?: number;
  readonly runBorder?: RetainedRunBorderFacts;
  readonly runBorderFragments?: readonly BorderSegment[];
  readonly revision?: Readonly<{ kind: string; author?: string }>;
  readonly typography?: RetainedRunTypographyFacts;
  readonly unsupportedGeometry?: readonly (
    | 'underline'
    | 'strikethrough'
    | 'double-strikethrough'
    | 'emphasis'
  )[];
  readonly decorations: readonly TextDecorationLayout[];
  readonly hyperlink?: string;
  readonly bookmark?: string;
}

export interface TabPlacement {
  readonly kind: 'tab';
  readonly range: TextRange;
  readonly bounds?: LayoutRect;
  readonly advancePt: number;
  readonly leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
  /** Fully repeated and positioned during acquisition; paint never measures. */
  readonly leaderGlyphs?: readonly RetainedGlyphPaintOperation[];
}

export interface AnchorHostPlacement {
  readonly kind: 'anchor-host';
  readonly range: TextRange;
  readonly bounds: LayoutRect;
  readonly baselinePt: number;
  readonly sourceMetrics?: Readonly<{ ascentPt: number; descentPt: number }>;
  readonly anchorOccurrenceId?: string;
}

export type InlineResourceKind = 'image' | 'chart' | 'math' | 'picture-bullet';
export type PaintResourceKind = InlineResourceKind;

export type PaintResourceDescriptorKind = InlineResourceKind;

export type ImagePaintResourceDescriptor = Readonly<{
  kind: 'image' | 'picture-bullet';
  resourceKey: string;
  partPath: string;
  mimeType: string;
  intrinsicSize: Readonly<{ widthPt: number; heightPt: number }>;
  svgImagePath?: string;
  srcRect?: Readonly<{ l: number; t: number; r: number; b: number }>;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  alpha?: number;
  colorReplaceFrom?: string;
  duotone?: DeepReadonly<Duotone>;
}>;

export type ChartPaintResourceDescriptor = Readonly<{
  kind: 'chart';
  resourceKey: string;
  intrinsicSize: Readonly<{ widthPt: number; heightPt: number }>;
  model: DeepReadonly<ChartModel>;
}>;

export type MathPaintResourceDescriptor = Readonly<{
  kind: 'math';
  resourceKey: string;
}>;

export type PaintResourceDescriptor =
  | ImagePaintResourceDescriptor
  | ChartPaintResourceDescriptor
  | MathPaintResourceDescriptor;

export interface PaintResourceRegistry {
  readonly keys: readonly string[];
  readonly descriptors: readonly DeepReadonly<PaintResourceDescriptor>[];
  resolve<K extends PaintResourceDescriptorKind>(
    resourceKey: string,
    expectedKind: K,
  ): DeepReadonly<Extract<PaintResourceDescriptor, { kind: K }>>;
}

export interface ResourcePlacement {
  readonly kind: 'resource';
  readonly range: TextRange;
  readonly resourceKey: string;
  readonly resourceKind: InlineResourceKind;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
}

export interface DrawingPlacement {
  readonly kind: 'drawing';
  readonly range: TextRange;
  readonly drawingId: LayoutNodeId;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
}

export type ParagraphPlacement =
  | TextPlacement
  | TabPlacement
  | AnchorHostPlacement
  | ResourcePlacement
  | DrawingPlacement;

export interface LineLayout {
  readonly range: TextRange;
  readonly bounds: LayoutRect;
  readonly baselinePt: number;
  readonly advancePt: number;
  readonly placements: readonly ParagraphPlacement[];
}

export type InlineResourceLayout = Readonly<{
  kind: InlineResourceKind;
  resourceKey: string;
  intrinsicSize: Readonly<{ widthPt: number; heightPt: number }>;
}>;

export interface BorderSegment {
  readonly edge?: 'top' | 'right' | 'bottom' | 'left' | 'between';
  readonly from: PointPt;
  readonly to: PointPt;
  readonly color: string;
  readonly widthPt: number;
  /** Exact authored ST_Border token. Kept independently of paint normalization. */
  readonly authoredStyle: string;
  readonly style: 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy';
  /** Final ST_Border cadence in point-space; empty for continuous/double rails. */
  readonly dashPatternPt?: readonly number[];
}

export type FillPaint = Readonly<{ color: string }>;

export interface WrapExclusion {
  readonly id: string;
  readonly wrap: 'square' | 'tight' | 'through' | 'topAndBottom';
  readonly bounds: LayoutRect;
  readonly polygon: readonly PointPt[];
  readonly anchorOccurrenceId?: string;
  readonly verticalOwnership?: 'page' | 'host';
}

export interface ParagraphFlowEvent {
  readonly kind: 'break';
  readonly breakKind: 'line' | 'page' | 'column';
  readonly offset: number;
}

/** Plain frame placement geometry consumed by layout and renderer adapters. */
export interface FrameGeometryState {
  readonly scale: number;
  readonly contentX: number;
  readonly contentW: number;
  readonly pageWidth: number;
  readonly pageH: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly y: number;
}

export interface ParagraphMarkLayout {
  readonly hidden: boolean;
  readonly bounds: LayoutRect;
}

export interface LineNumberPaintOperation {
  readonly kind: 'text';
  readonly text: string;
  readonly origin: PointPt;
  readonly font: string;
  readonly color: string;
  readonly textAlign: 'right';
}

/** ECMA-376 §17.6.8 retained line counter and its optional paint operation. */
export interface LineNumberLayout {
  readonly lineIndex: number;
  readonly counterValue: number;
  readonly bounds: LayoutRect;
  readonly paintOps: readonly LineNumberPaintOperation[];
}

export interface ParagraphSpacingLayout {
  readonly beforePt: number;
  readonly afterPt: number;
}

export interface ParagraphLayout extends LayoutNodeBase {
  readonly kind: 'paragraph';
  readonly styleId?: string | null;
  readonly spacing: ParagraphSpacingLayout;
  readonly contextualSpacing: boolean;
  readonly lines: readonly LineLayout[];
  readonly borders: readonly BorderSegment[];
  readonly shading?: FillPaint;
  readonly resources: readonly InlineResourceLayout[];
  readonly drawings: readonly DrawingLayout[];
  readonly textBoxes: readonly TextBoxLayout[];
  readonly events: readonly ParagraphFlowEvent[];
  readonly exclusions: readonly WrapExclusion[];
  readonly anchorFrames?: readonly AnchorFrameResult[];
  readonly paragraphMark?: ParagraphMarkLayout;
  readonly lineNumbers?: readonly LineNumberLayout[];
  readonly continuation?: Readonly<{
    lineStart: number;
    lineEnd: number;
    continuesFromPrevious: boolean;
    continuesOnNext: boolean;
  }>;
}

export interface TableLayout extends LayoutNodeBase {
  readonly kind: 'table';
}

export interface TextBoxLayout extends LayoutNodeBase {
  readonly kind: 'textbox';
  readonly paragraphs: readonly ParagraphLayout[];
  readonly writingMode: WritingMode;
  readonly verticalMode?: 'vert' | 'vert270' | 'eaVert' | 'mongolianVert';
  readonly contentBounds?: LayoutRect;
  readonly insets: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>;
}

export interface NoteLayout extends LayoutNodeBase {
  readonly kind: 'note';
}

export type PaintNode = ParagraphLayout | TableLayout | DrawingLayout | TextBoxLayout | NoteLayout;

export type PageLayerId =
  | 'background'
  | 'behindText'
  | 'header'
  | 'body'
  | 'notes'
  | 'front'
  | 'footer';

export interface PagePaintEntry {
  readonly layer: PageLayerId;
  readonly nodeId: LayoutNodeId;
}

export interface PageLayers {
  readonly paintOrder: readonly PagePaintEntry[];
  readonly background: readonly PaintNode[];
  readonly behindText: readonly PaintNode[];
  readonly header: readonly PaintNode[];
  readonly body: readonly PaintNode[];
  readonly notes: readonly PaintNode[];
  readonly front: readonly PaintNode[];
  readonly footer: readonly PaintNode[];
}

export interface LayoutPage {
  readonly pageIndex: number;
  readonly geometry: PageGeometry;
  readonly flowDomains: readonly FlowDomain[];
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly layers: PageLayers;
  readonly readingOrder: readonly LayoutNodeId[];
}

export type LayoutDiagnosticCode =
  | 'FLOW_OVERLAP'
  | 'BOTTOM_MARGIN_INVASION'
  | 'FLOW_DOMAIN_INVASION'
  | 'INVALID_REFERENCE'
  | 'INVALID_GEOMETRY'
  | 'NON_CONVERGENCE'
  | 'UNSUPPORTED_FEATURE';

export interface LayoutDiagnostic {
  readonly code: LayoutDiagnosticCode;
  readonly severity: 'warning' | 'error';
  readonly source?: SourceRef;
  readonly message: string;
}

export interface DocumentLayout {
  readonly pages: readonly LayoutPage[];
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export type CompatibilityEvidence =
  | Readonly<{ kind: 'microsoft-note'; reference: string }>
  | Readonly<{
      kind: 'office-observation';
      syntheticFixtureId: string;
      application: string;
      version: string;
      platform: string;
    }>;

export interface CompatibilityRule {
  readonly id: string;
  readonly evidence: CompatibilityEvidence;
  readonly description: string;
}

export interface LayoutServices {
  readonly text: TextLayoutService;
  readonly images: ImageMetadataService;
  readonly math: MathMetadataService;
}

/** Plain, parser-independent input for shaping a numbering marker. The renderer
 * boundary snapshots effective level rPr facts into this contract before the
 * retained layout service sees them. */
export interface NumberingMarkerShapeInput {
  readonly fontSizePt: number;
  readonly fonts: TextFontSlots;
  readonly themeFonts?: TextFontSlots;
  readonly themeFontPresence?: TextFontSlotPresence;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
  readonly complexScript: boolean;
  readonly fontHint?: 'default' | 'eastAsia' | 'cs';
  readonly eastAsiaLanguage?: string;
  readonly kerning?: boolean;
}

export interface ParagraphLayoutInput {
  readonly kind: 'paragraph';
  readonly source: SourceRef;
}

export interface AcquiredParagraphLayoutInput {
  readonly kind: 'paragraph';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly ordinaryFlow: boolean;
  readonly styleId?: string | null;
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly spacing: ParagraphSpacingLayout;
  readonly contextualSpacing?: boolean;
  readonly lines: readonly LineLayout[];
  readonly borders: readonly BorderSegment[];
  readonly shading?: FillPaint;
  readonly resources: readonly InlineResourceLayout[];
  readonly drawings: readonly DrawingLayout[];
  readonly textBoxes: readonly TextBoxLayout[];
  readonly events: readonly ParagraphFlowEvent[];
  readonly exclusions: readonly WrapExclusion[];
  readonly anchorFrames?: readonly AnchorFrameResult[];
  readonly paragraphMark?: ParagraphMarkLayout;
  readonly continuation?: Readonly<{
    lineStart: number;
    lineEnd: number;
    continuesFromPrevious: boolean;
    continuesOnNext: boolean;
  }>;
}

export interface TableLayoutInput {
  readonly kind: 'table';
  readonly source: SourceRef;
}

export type FlowBlockInput = ParagraphLayoutInput | TableLayoutInput;

export interface FlowContainer extends FlowDomain {}

export interface FlowCursor extends PointPt {}

export interface FlowBlockPlacement {
  readonly container: FlowContainer;
  readonly cursor: FlowCursor;
  readonly availableBounds: LayoutRect;
}

export interface BlockLayoutResult<T extends ParagraphLayout | TableLayout> {
  readonly layout: T;
  readonly nextCursor: FlowCursor;
}

export interface FlowLayoutInput {
  readonly blocks: readonly FlowBlockInput[];
  readonly container: FlowContainer;
  readonly cursor: FlowCursor;
  readonly source: SourceRef;
}

export interface FlowLayout extends FlowOwnership {
  readonly source: SourceRef;
  readonly container: FlowContainer;
  readonly blocks: readonly (ParagraphLayout | TableLayout)[];
  readonly nextCursor: FlowCursor;
}

export interface BlockLayoutAlgorithms {
  layoutParagraph(
    input: ParagraphLayoutInput,
    placement: FlowBlockPlacement,
    services: LayoutServices,
  ): BlockLayoutResult<ParagraphLayout>;
  layoutTable(
    input: TableLayoutInput,
    placement: FlowBlockPlacement,
    services: LayoutServices,
  ): BlockLayoutResult<TableLayout>;
}
