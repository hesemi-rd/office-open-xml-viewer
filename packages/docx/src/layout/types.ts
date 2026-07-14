import type { SectionLayoutContext } from '../layout-context.js';

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

export type DrawingPaintCommand = Readonly<{
  kind: 'fill-rect';
  rect: LayoutRect;
  fill: string;
}>;

export interface DrawingLayout extends LayoutNodeBase {
  readonly kind: 'drawing';
  readonly transform?: Matrix2DData;
  readonly clip?: ClipPathData;
  readonly commands: readonly DrawingPaintCommand[];
}

export interface ParagraphLayout extends LayoutNodeBase {
  readonly kind: 'paragraph';
}

export interface TableLayout extends LayoutNodeBase {
  readonly kind: 'table';
}

export interface TextBoxLayout extends LayoutNodeBase {
  readonly kind: 'textbox';
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

export interface TextLayoutService {
  readonly fingerprint: string;
}

export interface ImageMetadataService {
  readonly fingerprint: string;
}

export interface MathMetadataService {
  readonly fingerprint: string;
}

export interface LayoutServices {
  readonly text: TextLayoutService;
  readonly images: ImageMetadataService;
  readonly math: MathMetadataService;
}

export interface ParagraphLayoutInput {
  readonly kind: 'paragraph';
  readonly source: SourceRef;
}

export interface TableLayoutInput {
  readonly kind: 'table';
  readonly source: SourceRef;
}

export type FlowBlockInput = ParagraphLayoutInput | TableLayoutInput;

export interface FlowContainer extends FlowDomain {}

export interface FlowLayoutInput {
  readonly blocks: readonly FlowBlockInput[];
  readonly container: FlowContainer;
  readonly source: SourceRef;
}

export interface FlowLayout extends FlowOwnership {
  readonly source: SourceRef;
  readonly container: FlowContainer;
  readonly blocks: readonly (ParagraphLayout | TableLayout)[];
}

export interface BlockLayoutAlgorithms {
  layoutParagraph(input: ParagraphLayoutInput, services: LayoutServices): ParagraphLayout;
  layoutTable(input: TableLayoutInput, services: LayoutServices): TableLayout;
}
