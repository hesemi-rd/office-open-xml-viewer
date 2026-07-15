import { describe, it, expect, beforeAll } from 'vitest';
import { bodyFragmentFor, paginateDocument, renderDocumentToCanvas } from './renderer.js';
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
import type { TableFragmentLayout } from './layout/table-pagination.js';

// ─────────────────────────────────────────────────────────────────────────────
// A5 represents split rows with TableFragmentLayout metadata instead of runtime
// `pageCutBottom` flags. A partial row always terminates its page fragment, so an
// intra-row page cut can never be reclassified as an interior row boundary. The
// fragment's outer-bottom border remains paintable at the page edge.
// ─────────────────────────────────────────────────────────────────────────────

interface StrokeSeg { x1: number; y1: number; x2: number; y2: number; width: number; color: string }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; strokes: StrokeSeg[]; measured: () => number } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const strokes: StrokeSeg[] = [];
  let cur = { x: 0, y: 0 };
  let pending: { x1: number; y1: number; x2: number; y2: number } | null = null;
  let lineWidth = 1;
  let strokeStyle = '#000000';
  let fillStyle = '#000000';
  let measured = 0;
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
      measured += 1;
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
  return { canvas: canvas as unknown as HTMLCanvasElement, strokes, measured: () => measured };
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
  const { canvas, strokes, measured } = makeRecordingCanvas();
  await renderDocumentToCanvas(model, canvas, pageIndex, { dpr: 1, width: 160, prebuiltPages: pages });
  expect(measured()).toBe(0);
  return strokes;
}

const horizontals = (strokes: StrokeSeg[]) =>
  strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.x2 - s.x1) > 10);

/** Find a retained fragment whose final logical row continues on a later page. */
function findSamePageCutPage(pages: PaginatedBodyElement[][]): {
  pageIndex: number;
  colTopPt: number;
  rowHeightsPt: number[];
  cutFlags: boolean[];
} | null {
  const retained = pages.map((page) => page.flatMap((element) => {
    if (element.type !== 'table') return [];
    const placed = bodyFragmentFor(element);
    if (placed?.fragment.kind !== 'table' || !('flowBounds' in placed.fragment)) {
      throw new Error('expected retained TableFragmentLayout');
    }
    expect(element).not.toHaveProperty('tableColWidthsPt');
    expect(element).not.toHaveProperty('tableRowHeightsPt');
    return [{ placed, fragment: placed.fragment as TableFragmentLayout }];
  }));
  for (let pi = 0; pi < pages.length; pi++) {
    const entry = retained[pi]?.[0];
    if (!entry) continue;
    const cutFlags = entry.fragment.rows.map((tableRow) => retained.slice(pi + 1).some((page) =>
      page.some(({ fragment }) => fragment.rows.some((laterRow) =>
        laterRow.logicalRowIndex === tableRow.logicalRowIndex
        && laterRow.fragmentIndex > tableRow.fragmentIndex))));
    if (cutFlags.some(Boolean)) {
      return {
        pageIndex: pi,
        colTopPt: entry.placed.yPt,
        rowHeightsPt: entry.fragment.rows.map((tableRow) => tableRow.advancePt),
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
  it('keeps an intra-row cut at the retained fragment boundary and paints measure-free', async () => {
    const model = twoTallRowsDoc(120);
    const pages = paginateDocument(model);
    const hit = findSamePageCutPage(pages);

    // The old same-page runtime-marker state is obsolete: retained pagination
    // guarantees that a continuing row is the final row in its page fragment.
    expect(hit).not.toBeNull();
    const { pageIndex, colTopPt, rowHeightsPt, cutFlags } = hit!;
    expect(cutFlags.slice(0, -1)).not.toContain(true);
    expect(cutFlags.at(-1)).toBe(true);
    expect(rowHeightsPt.length).toBeGreaterThanOrEqual(1);

    const topY = colTopPt;
    const pageEndY = colTopPt + rowHeightsPt.reduce((s, v) => s + v, 0);

    const h = horizontals(await renderPage(model, pages, pageIndex));

    // The table's own OUTER top edge (slice top) is still drawn.
    expect(nearFullWidth(h, topY, 160).length).toBeGreaterThanOrEqual(1);
    // The page cut is an outer-bottom boundary, never an insideH conflict.
    expect(nearFullWidth(h, pageEndY, 160).length).toBeGreaterThanOrEqual(1);
  });
});
