import { describe, it, expect, beforeAll } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableRow,
  DocxDocumentModel,
  PaginatedBodyElement,
  SectionProps,
  BorderSpec,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Intra-row page-cut border on a SAME-PAGE continuation piece. A tall table row
// that the paginator splits into pieces can land TWO of those pieces on the SAME
// page (measured private fixture sample-33 p.3: two consecutive tall source rows
// each split, a continuation piece of each sharing one page — both pieces carry
// the runtime `pageCutBottom` marker). The leading piece's bottom is then an
// INTERIOR horizontal edge (not the table's outer bottom), so drawTableRows
// resolved it against the piece below via §17.4.66 and drew the Table-Grid
// insideH — a full-width rule Word does not draw at an intra-row cut. Word leaves
// that cut OPEN. Only the true page-end cut (the last piece on the page, resolved
// in the outer-bottom branch) keeps its rule.
// ─────────────────────────────────────────────────────────────────────────────

interface StrokeSeg { x1: number; y1: number; x2: number; y2: number; width: number; color: string }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; strokes: StrokeSeg[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const strokes: StrokeSeg[] = [];
  let cur = { x: 0, y: 0 };
  let pending: { x1: number; y1: number; x2: number; y2: number } | null = null;
  let lineWidth = 1;
  let strokeStyle = '#000000';
  let fillStyle = '#000000';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get lineWidth() { return lineWidth; },
    set lineWidth(v: number) { lineWidth = v; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(v: string) { strokeStyle = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    textAlign: 'left' as CanvasTextAlign,
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() { pending = null; }, closePath() {},
    moveTo(x: number, y: number) { cur = { x, y }; },
    lineTo(x: number, y: number) { pending = { x1: cur.x, y1: cur.y, x2: x, y2: y }; cur = { x, y }; },
    stroke() {
      if (pending) strokes.push({ ...pending, width: lineWidth, color: strokeStyle });
    },
    fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, strokes };
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
  };
});

const bs = (): BorderSpec => ({ style: 'single', width: 1, color: null } as BorderSpec);
const allEdges = () => ({ top: bs(), bottom: bs(), left: bs(), right: bs(), insideH: bs(), insideV: bs() });

function para(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 20, color: null, fontFamily: 'T',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 20, defaultFontFamily: 'T', widowControl: false,
  } as unknown as DocParagraph;
}

/** A tall single-cell row built from `nParas` single-line paragraphs. Many short
 *  blocks force the height-splitter down its block path (splitRowByCellBlocks),
 *  whose whole-block pieces pack multiple continuation pieces onto one page. */
function tallBlockRow(nParas: number): DocTableRow {
  const content: CellElement[] = [];
  for (let i = 0; i < nParas; i++) {
    content.push({ type: 'paragraph', ...para(`P${i}`) } as unknown as CellElement);
  }
  return {
    cells: [{
      content, colSpan: 1, vMerge: null, borders: allEdges(),
      background: null, vAlign: 'top', widthPt: null,
    }],
    rowHeight: null, rowHeightRule: 'auto', isHeader: false,
  } as unknown as DocTableRow;
}

function twoTallRowsDoc(pageHeight: number): DocxDocumentModel {
  const t: DocTable = {
    colWidths: [160], rows: [tallBlockRow(10), tallBlockRow(10)], borders: allEdges(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
  return {
    section: {
      pageWidth: 160, pageHeight,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...t } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { T: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderPage(model: DocxDocumentModel, pages: PaginatedBodyElement[][], pageIndex: number): Promise<StrokeSeg[]> {
  const { canvas, strokes } = makeRecordingCanvas();
  await renderDocumentToCanvas(model, canvas, pageIndex, { dpr: 1, width: 160, prebuiltPages: pages });
  return strokes;
}

const horizontals = (strokes: StrokeSeg[]) =>
  strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.x2 - s.x1) > 10);

/** Find the page whose table carries two same-source-row split pieces, i.e. a
 *  NON-LAST row on the page marked pageCutBottom. Returns page index + geometry. */
function findSamePageCutPage(pages: PaginatedBodyElement[][]): {
  pageIndex: number;
  colTopPt: number;
  rowHeightsPt: number[];
  cutFlags: boolean[];
} | null {
  for (let pi = 0; pi < pages.length; pi++) {
    const tbl = pages[pi].find((el) => (el as { type?: string }).type === 'table') as
      (PaginatedBodyElement & DocTable & { colTopPt?: number; tableRowHeightsPt?: number[] }) | undefined;
    if (!tbl || !tbl.rows) continue;
    const cutFlags = tbl.rows.map((r) => (r as DocTableRow & { pageCutBottom?: boolean }).pageCutBottom === true);
    if (tbl.rows.length >= 2 && cutFlags.slice(0, -1).some(Boolean)) {
      return {
        pageIndex: pi,
        colTopPt: tbl.colTopPt ?? 0,
        rowHeightsPt: tbl.tableRowHeightsPt ?? [],
        cutFlags,
      };
    }
  }
  return null;
}

/** Full-width horizontals within ±tol of `y`. */
const nearFullWidth = (h: StrokeSeg[], y: number, pageW: number, tol = 1.2) =>
  h.filter((s) =>
    Math.abs(s.y1 - y) <= tol &&
    Math.min(s.x1, s.x2) <= 1 &&
    Math.max(s.x1, s.x2) >= pageW - 1);

describe('#986 same-page intra-row cut border', () => {
  it('draws NO rule at an intra-row cut whose continuation piece shares the page', async () => {
    // Build the renderer-facing page slice directly. Pagination geometry is
    // covered separately; this regression is specifically the paint contract
    // for two row pieces sharing a page, with the leading piece ending at an
    // intra-source-row cut.
    const model = twoTallRowsDoc(120);
    const sourceTable = model.body[0] as unknown as DocTable;
    const leading = tallBlockRow(2) as DocTableRow & { pageCutBottom?: boolean };
    leading.pageCutBottom = true;
    const slice = {
      ...sourceTable,
      type: 'table',
      rows: [leading, tallBlockRow(2)],
      colTopPt: 0,
      tableColWidthsPt: [160],
      tableRowHeightsPt: [41, 41],
      tableContentWPt: 160,
    } as unknown as PaginatedBodyElement;
    const pages: PaginatedBodyElement[][] = [[slice]];
    const hit = findSamePageCutPage(pages);

    // Fixture sanity: the same-page two-piece condition really exists and the
    // LEADING (non-last) piece is the intra-row cut.
    expect(hit).not.toBeNull();
    const { pageIndex, colTopPt, rowHeightsPt, cutFlags } = hit!;
    expect(cutFlags[0]).toBe(true);
    expect(rowHeightsPt.length).toBeGreaterThanOrEqual(2);

    const topY = colTopPt;
    const interiorY = colTopPt + rowHeightsPt[0];
    const pageEndY = colTopPt + rowHeightsPt.reduce((s, v) => s + v, 0);

    const h = horizontals(await renderPage(model, pages, pageIndex));

    // The table's own OUTER top edge (slice top) is still drawn.
    expect(nearFullWidth(h, topY, 160).length).toBeGreaterThanOrEqual(1);
    // The true page-end cut (the LAST piece's outer bottom) is still drawn.
    expect(nearFullWidth(h, pageEndY, 160).length).toBeGreaterThanOrEqual(1);
    // …but the INTERIOR intra-row cut draws NOTHING (Word leaves it open). This
    // is the regression assertion: it FAILS before the fix (a Table-Grid insideH
    // rule is drawn at interiorY) and PASSES once the else-branch guard suppresses
    // the interior edge of a pageCutBottom-marked continuation piece.
    expect(nearFullWidth(h, interiorY, 160)).toHaveLength(0);
  });
});
