import { DocxDocument } from './document';
import type { Document } from './types';

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
type DiffFn = (a: Document, b: Document) => DiffResult;

export interface DocxDiffViewerOptions {
  width?: number;
  dpr?: number;
  useGoogleFonts?: boolean;
  onPageChange?: (index: number, total: number) => void;
  onDiff?: (result: DiffResult) => void;
}

/**
 * Side-by-side DOCX diff viewer.
 *
 * Renders two documents into two canvases (one page each) and reports the
 * structural diff via `onDiff`. The diff result includes per-paragraph
 * locations; consumers can render their own changes list alongside.
 *
 * Inject the diff implementation via `setDiffFn` (this keeps the docx package
 * independent of `@silurus/ooxml-diff` at build time).
 */
export class DocxDiffViewer {
  private readonly left: HTMLCanvasElement;
  private readonly right: HTMLCanvasElement;
  private leftDoc: DocxDocument | null = null;
  private rightDoc: DocxDocument | null = null;
  private readonly opts: DocxDiffViewerOptions;
  private diffFn: DiffFn | null = null;
  private result: DiffResult | null = null;
  private currentPage = 0;

  constructor(left: HTMLCanvasElement, right: HTMLCanvasElement, opts: DocxDiffViewerOptions = {}) {
    this.opts = opts;
    this.left = left;
    this.right = right;
  }

  setDiffFn(fn: (a: Document, b: Document) => unknown): void {
    this.diffFn = fn as DiffFn;
  }

  async load(before: string | ArrayBuffer, after: string | ArrayBuffer): Promise<void> {
    this.leftDoc = await DocxDocument.load(before, { useGoogleFonts: this.opts.useGoogleFonts });
    this.rightDoc = await DocxDocument.load(after, { useGoogleFonts: this.opts.useGoogleFonts });
    this._computeDiff();
    this.currentPage = 0;
    await this._renderBoth();
  }

  get pageCount(): number {
    return Math.max(this.leftDoc?.pageCount ?? 0, this.rightDoc?.pageCount ?? 0);
  }

  get currentPageIndex(): number { return this.currentPage; }

  get diffResult(): DiffResult | null { return this.result; }

  async goToPage(index: number): Promise<void> {
    if (this.pageCount === 0) return;
    this.currentPage = Math.max(0, Math.min(index, this.pageCount - 1));
    await this._renderBoth();
  }

  async nextPage(): Promise<void> { await this.goToPage(this.currentPage + 1); }
  async prevPage(): Promise<void> { await this.goToPage(this.currentPage - 1); }

  private _computeDiff(): void {
    if (!this.leftDoc || !this.rightDoc || !this.diffFn) {
      this.result = null;
      return;
    }
    this.result = this.diffFn(this.leftDoc.document, this.rightDoc.document);
    this.opts.onDiff?.(this.result);
  }

  private async _renderBoth(): Promise<void> {
    const width = this.opts.width ?? this.left.offsetWidth ?? 600;
    const dpr = this.opts.dpr ?? (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);

    if (this.leftDoc && this.currentPage < this.leftDoc.pageCount) {
      await this.leftDoc.renderPage(this.left, this.currentPage, { width, dpr });
    } else {
      this._clear(this.left);
    }
    if (this.rightDoc && this.currentPage < this.rightDoc.pageCount) {
      await this.rightDoc.renderPage(this.right, this.currentPage, { width, dpr });
    } else {
      this._clear(this.right);
    }

    this.opts.onPageChange?.(this.currentPage, this.pageCount);
  }

  private _clear(canvas: HTMLCanvasElement): void {
    const ctx = canvas.getContext('2d');
    if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
  }

  destroy(): void {
    this.leftDoc?.destroy();
    this.rightDoc?.destroy();
  }
}
