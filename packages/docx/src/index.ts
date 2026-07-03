export { DocxDocument, type LoadOptions } from './document';
export type { WireRenderPageOptions } from './worker-protocol';
export { DocxViewer, type DocxViewerOptions } from './viewer';
export { DocxScrollViewer, type DocxScrollViewerOptions } from './scroll-viewer';
export { buildDocxTextLayer } from './text-layer';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
// Typed load-time error surfaced by DocxDocument.load (e.g. a password-protected
// or legacy-binary .doc file). Re-exported so `@silurus/ooxml/docx` consumers can
// narrow on `err.code`.
export { OoxmlError, type OoxmlErrorCode } from '@silurus/ooxml-core';
export { noteText } from './types';
export type {
  DocxDocumentModel,
  DocSettings,
  SectionProps,
  // Per-section page geometry (reachable via the BodyElement sectionBreak arm's
  // `geom` and PaginatedBodyElement's `sectionGeom`, ECMA-376 §17.6.13/§17.6.11).
  SectionGeom,
  // Multi-column section sub-types (reachable via SectionProps.columns).
  ColumnsSpec,
  ColSpec,
  HeadersFooters,
  HeaderFooter,
  NumberingInfo,
  BodyElement,
  DocParagraph,
  DocRun,
  // Absolute-position tab run (reachable via the DocRun union's `ptab` arm,
  // ECMA-376 §17.3.3.23).
  PTabRun,
  DocxTextRun,
  FieldRun,
  ImageRun,
  ShapeRun,
  ShapeText,
  // Per-run shape-text formatting (reachable via ShapeText.runs).
  ShapeTextRun,
  RubyAnnotation,
  RenderPageOptions,
  RunRevision,
  DocRevision,
  DocComment,
  DocNote,
  NoteRef,
  // Paragraph / line-spacing sub-types.
  LineSpacing,
  // Text-frame / drop-cap properties (reachable via DocParagraph.framePr).
  FramePr,
  TabStop,
  ParagraphBorders,
  ParaBorderEdge,
  DocxRunBorder,
  // Table model (reachable via BodyElement table variant).
  DocTable,
  TblpPr,
  DocTableRow,
  DocTableCell,
  CellElement,
  TableBorders,
  CellBorders,
  BorderSpec,
  // Shape geometry / fill sub-types.
  PathCmd,
  GradientStop,
  LineEnd,
} from './types';
export type { DocxTextRunInfo } from './renderer';
