import { XlsxWorkbook } from './workbook';
import type { Worksheet, Workbook } from './types';
import { XlsxViewer } from './viewer';

interface DiffInput { workbook: Workbook; worksheets: Record<string, Worksheet>; }

interface BBox { x: number; y: number; width: number; height: number; }
type ChangeLocation =
  | { kind: 'slide'; slideIndex: number; bbox?: BBox }
  | { kind: 'page'; pageIndex: number; bbox?: BBox }
  | { kind: 'paragraph'; paragraphIndex: number }
  | { kind: 'table'; row: number; col: number }
  | { kind: 'cell'; sheetName: string; row: number; col: number }
  | { kind: 'sheet'; sheetName: string };
type DiffChange = {
  op: 'add' | 'remove' | 'modify';
  path: string;
  kind: string;
  before?: unknown;
  after?: unknown;
  location?: ChangeLocation;
};
type DiffResult = { format: 'pptx' | 'docx' | 'xlsx'; changes: DiffChange[] };
type DiffFn = (a: DiffInput, b: DiffInput) => DiffResult;

export interface XlsxDiffViewerOptions {
  cellScale?: number;
  useGoogleFonts?: boolean;
  onDiff?: (result: DiffResult) => void;
}

/**
 * Side-by-side XLSX diff viewer.
 *
 * Hosts two {@link XlsxViewer} instances inside the caller-supplied containers
 * and reports a structural diff (cells, merges, sheet metadata) via `onDiff`.
 *
 * Inject the diff implementation via `setDiffFn` (the xlsx package stays
 * independent of `@silurus/ooxml-diff` at build time).
 */
export class XlsxDiffViewer {
  private readonly leftViewer: XlsxViewer;
  private readonly rightViewer: XlsxViewer;
  private leftWb: XlsxWorkbook | null = null;
  private rightWb: XlsxWorkbook | null = null;
  private readonly opts: XlsxDiffViewerOptions;
  private diffFn: DiffFn | null = null;
  private result: DiffResult | null = null;

  constructor(leftContainer: HTMLElement, rightContainer: HTMLElement, opts: XlsxDiffViewerOptions = {}) {
    this.opts = opts;
    this.leftViewer = new XlsxViewer(leftContainer, { cellScale: opts.cellScale, useGoogleFonts: opts.useGoogleFonts });
    this.rightViewer = new XlsxViewer(rightContainer, { cellScale: opts.cellScale, useGoogleFonts: opts.useGoogleFonts });
  }

  setDiffFn(fn: (a: DiffInput, b: DiffInput) => unknown): void {
    this.diffFn = fn as DiffFn;
  }

  async load(before: string | ArrayBuffer, after: string | ArrayBuffer): Promise<void> {
    this.leftWb = new XlsxWorkbook();
    this.rightWb = new XlsxWorkbook();
    await Promise.all([
      this.leftWb.load(before, { useGoogleFonts: this.opts.useGoogleFonts }),
      this.rightWb.load(after, { useGoogleFonts: this.opts.useGoogleFonts }),
    ]);

    // Render each side's first sheet via the inner viewers. Each viewer owns
    // its own XlsxWorkbook; we just call its public load so it picks up the
    // tab bar / scroll host wiring.
    await Promise.all([
      this.leftViewer.load(before),
      this.rightViewer.load(after),
    ]);

    await this._computeDiff();
  }

  get diffResult(): DiffResult | null { return this.result; }

  private async _computeDiff(): Promise<void> {
    if (!this.leftWb || !this.rightWb || !this.diffFn) {
      this.result = null;
      return;
    }
    const a = this.leftWb.parsed;
    const b = this.rightWb.parsed;
    if (!a || !b) return;

    // Eagerly load every sheet on both sides so the diff sees full cell data.
    const aSheets: Record<string, Worksheet> = {};
    const bSheets: Record<string, Worksheet> = {};
    for (let i = 0; i < a.workbook.sheets.length; i++) {
      aSheets[a.workbook.sheets[i].name] = await this.leftWb.getWorksheet(i);
    }
    for (let i = 0; i < b.workbook.sheets.length; i++) {
      bSheets[b.workbook.sheets[i].name] = await this.rightWb.getWorksheet(i);
    }

    this.result = this.diffFn(
      { workbook: a.workbook, worksheets: aSheets },
      { workbook: b.workbook, worksheets: bSheets },
    );
    this.opts.onDiff?.(this.result);
  }

  destroy(): void {
    // XlsxViewer doesn't expose a destroy at the time of writing; the host
    // can remove the container to clean up.
  }
}
