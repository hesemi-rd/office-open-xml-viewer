/** Structural diff types shared across pptx / docx / xlsx producers. */

export type ChangeOp = 'add' | 'remove' | 'modify';

export interface BBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

export type ChangeLocation =
  | { kind: 'slide'; slideIndex: number; bbox?: BBox }
  | { kind: 'page'; pageIndex: number; bbox?: BBox }
  | { kind: 'paragraph'; paragraphIndex: number }
  | { kind: 'table'; row: number; col: number }
  | { kind: 'cell'; sheetName: string; row: number; col: number }
  | { kind: 'sheet'; sheetName: string };

export interface Change {
  op: ChangeOp;
  /** Dotted path inside the document model. Example: `slides[2].elements[5].textBody`. */
  path: string;
  /** Coarse category for filtering — `slide`, `element`, `text`, `geometry`, `fill`,
   *  `stroke`, `cell`, `paragraph`, `run`, `table-cell`, `image`, … */
  kind: string;
  before?: unknown;
  after?: unknown;
  location?: ChangeLocation;
}

export type Format = 'pptx' | 'docx' | 'xlsx';

export interface DiffResult {
  format: Format;
  changes: Change[];
}

/** Filter helper — keep only changes whose location matches a predicate. */
export function changesAtSlide(result: DiffResult, slideIndex: number): Change[] {
  return result.changes.filter(
    (c) => c.location?.kind === 'slide' && c.location.slideIndex === slideIndex,
  );
}

export function changesAtPage(result: DiffResult, pageIndex: number): Change[] {
  return result.changes.filter(
    (c) => c.location?.kind === 'page' && c.location.pageIndex === pageIndex,
  );
}

export function changesAtSheet(result: DiffResult, sheetName: string): Change[] {
  return result.changes.filter(
    (c) =>
      (c.location?.kind === 'sheet' || c.location?.kind === 'cell') &&
      'sheetName' in c.location &&
      c.location.sheetName === sheetName,
  );
}
