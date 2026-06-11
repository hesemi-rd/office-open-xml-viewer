export { DocxDocument, type LoadOptions } from './document';
export { DocxViewer, type DocxViewerOptions } from './viewer';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
export type {
  DocxDocumentModel,
  DocSettings,
  SectionProps,
  HeadersFooters,
  HeaderFooter,
  NumberingInfo,
  BodyElement,
  DocParagraph,
  DocRun,
  DocxTextRun,
  FieldRun,
  ImageRun,
  ShapeRun,
  ShapeText,
  RubyAnnotation,
  RenderPageOptions,
  RunRevision,
  DocRevision,
  DocComment,
  DocNote,
  // Paragraph / line-spacing sub-types.
  LineSpacing,
  TabStop,
  ParagraphBorders,
  ParaBorderEdge,
  // Table model (reachable via BodyElement table variant).
  DocTable,
  DocTableRow,
  DocTableCell,
  CellElement,
  TableBorders,
  CellBorders,
  BorderSpec,
  // Shape geometry / fill sub-types.
  PathCmd,
  GradientStop,
} from './types';
export type { DocxTextRunInfo } from './renderer';
