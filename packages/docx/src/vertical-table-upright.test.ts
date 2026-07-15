import { describe, expect, it } from 'vitest';
import {
  bodyFragmentFor,
  computePages,
  renderDocumentToCanvas,
  __test_verticalLayoutSection,
  type DocxTextRunInfo,
} from './renderer.js';
import type {
  DocxDocumentModel,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocParagraph,
  SectionProps,
  BodyElement,
} from './types';

// ECMA-376 §17.6.20 + §17.4.80/§17.18.37 — issue #988 batch-3 adjudication ④:
// a table CELL inside a vertical (tbRl) section renders like a normal
// HORIZONTAL cell — the section's vertical text direction does NOT propagate
// into the cell:
//   - cell text is laid out horizontally (left→right, wrapping downward),
//   - a fixed `tcW` is the cell's PHYSICAL horizontal width,
//   - `trHeight hRule="exact"` clips overflow at the physical row height,
//   - auto row height GROWS to enclose the content.
// The table block sits upright at the flow position: its physical top edge is
// the top content margin (the column axis start) and it advances the vertical
// flow by its PHYSICAL WIDTH (columns progress right→left past it).
//
// Word ground truth = the batch-3 vertical-table fixture PDF (two 1-cell
// 1-in-wide fixed tables, exact 2 in vs auto): both cells lay their long CJK
// string out horizontally at the physical width; the exact row's border box is
// exactly 2 in tall with lines 10+ clipped; the auto row grew to fit all lines.

/** Recording 2D context (same skeleton as table-clip-exact.test.ts): captures
 *  `rect()` calls so the §17.4.80 exact-row clip band can be asserted, with
 *  deterministic char-count metrics. */
interface RectCall { x: number; y: number; w: number; h: number; }
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  rectCalls: RectCall[];
  measureCalls: () => number;
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const rectCalls: RectCall[] = [];
  let measures = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measures += 1;
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {},
    rect(x: number, y: number, w: number, h: number) {
      rectCalls.push({ x, y, w, h });
    },
    clip() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {},
    setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    rectCalls,
    measureCalls: () => measures,
  };
}

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}

function bodyParagraph(text: string): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

/** 1×1 fixed-width (50 pt) table. `exactPt` sets `trHeight hRule="exact"`;
 *  null keeps the row auto. The cell text ('tok tok …') wraps one token per
 *  line at the 50 pt cell width (each 'tok ' measures 40 pt with the fake
 *  10 px-per-char metrics), giving 12 lines ≈ 130+ pt of content — far taller
 *  than an 80 pt exact row. */
function fixedTable(text: string, exactPt: number | null): DocTable {
  const cell: DocTableCell = {
    content: [{ type: 'paragraph', ...bodyParagraph(text) }],
    colSpan: 1,
    vMerge: null,
    borders: emptyBorders(),
    background: null,
    vAlign: 'top',
    widthPt: 50,
    marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0,
  } as unknown as DocTableCell;
  const row: DocTableRow = {
    cells: [cell],
    rowHeight: exactPt,
    rowHeightRule: exactPt == null ? 'auto' : 'exact',
    isHeader: false,
  } as unknown as DocTableRow;
  return {
    colWidths: [50],
    rows: [row],
    borders: emptyBorders(),
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

// PHYSICAL portrait page 200×300 pt with DISTINCT margins so the logical
// (swapped) frame's coordinates can never coincide with the physical ones:
//   physical top=20 right=30 bottom=40 left=24
//   ⇒ logical  left=20 top=30   right=40  bottom=24 (verticalLayoutSection)
// Flow starts at logical y = 30; the column axis starts at logical x = 20
// (= the physical top margin). Flow budget = 200 − 30 − 24 = 146 pt.
const PHYS = {
  pageWidth: 200, pageHeight: 300,
  marginTop: 20, marginRight: 30, marginBottom: 40, marginLeft: 24,
  headerDistance: 0, footerDistance: 0,
  titlePage: false, evenAndOddHeaders: false,
  textDirection: 'tbRl',
} as unknown as SectionProps;

const CSS_W = 200; // physical page width == canvas CSS width at scale 1
const FLOW_TOP = 30; // logical body top (physical right margin)
const COL_TOP = 20; // physical column top (logical marginLeft)
const TABLE_W = 50; // fixed tcW ⇒ physical table width

const CELL_TEXT_1 = 'aaa '.repeat(12).trim();
const CELL_TEXT_2 = 'bbb '.repeat(12).trim();

function verticalTableDoc(): DocxDocumentModel {
  return {
    section: PHYS,
    body: [
      { type: 'table', ...fixedTable(CELL_TEXT_1, 80) },
      { type: 'paragraph', ...bodyParagraph('zz') },
      { type: 'table', ...fixedTable(CELL_TEXT_2, null) },
      { type: 'paragraph', ...bodyParagraph('qq') },
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderRuns(): Promise<{ runs: DocxTextRunInfo[]; rectCalls: RectCall[] }> {
  const { canvas, rectCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(verticalTableDoc(), canvas, 0, {
    dpr: 1,
    width: PHYS.pageWidth,
    onTextRun: (r) => runs.push(r),
  });
  return { runs, rectCalls };
}

describe('vertical (tbRl) table cells render upright/horizontal (§17.6.20 + §17.4.80, #988 ④)', () => {
  it('cell text is horizontal at the fixed physical width, from the physical column top', async () => {
    const { runs } = await renderRuns();
    const cellRuns = runs.filter((r) => r.text.startsWith('aaa'));
    expect(cellRuns.length).toBeGreaterThanOrEqual(2);
    for (const r of cellRuns) {
      // Horizontal: no +90° overlay rotation.
      expect(r.transform, `run ${JSON.stringify(r)}`).toBeUndefined();
      // Fixed tcW ⇒ the physical x band [cssW − flowTop − tableW, +tableW].
      expect(r.x).toBeGreaterThanOrEqual(CSS_W - FLOW_TOP - TABLE_W - 2);
      expect(r.x + r.w).toBeLessThanOrEqual(CSS_W - FLOW_TOP + 2);
    }
    // Lines share the same x (left-aligned) and stack DOWN the physical page
    // from the physical column top (top content margin).
    const ys = cellRuns.map((r) => r.y);
    expect(new Set(cellRuns.map((r) => Math.round(r.x))).size).toBe(1);
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(COL_TOP - 1);
    expect(Math.min(...ys)).toBeLessThanOrEqual(COL_TOP + 15);
    for (let i = 1; i < ys.length; i++) expect(ys[i]).toBeGreaterThan(ys[i - 1]);
  });

  it('trHeight exact clips at the physical row height; auto grows past it', async () => {
    const { runs, rectCalls } = await renderRuns();
    // §17.4.80 exact ⇒ an 80 pt clip band. The retained painter submits
    // point-space clip geometry before the canvas placement transform, while
    // the text-run assertions below observe the resulting physical placement.
    const clipRect = rectCalls.find(
      (r) => Math.abs(r.h - 80) < 1e-6,
    );
    expect(clipRect, 'exact row must clip at its physical Y band').toBeDefined();
    // auto ⇒ the row grows: content extends well past the 80 pt exact height.
    const autoRuns = runs.filter((r) => r.text.startsWith('bbb'));
    expect(autoRuns.length).toBeGreaterThanOrEqual(2);
    for (const r of autoRuns) expect(r.transform).toBeUndefined();
    expect(Math.max(...autoRuns.map((r) => r.y))).toBeGreaterThan(COL_TOP + 80);
  });

  it('the table advances the flow by its PHYSICAL WIDTH; body text stays vertical', async () => {
    const { runs } = await renderRuns();
    const afterExact = runs.find((r) => r.text === 'zz');
    expect(afterExact).toBeDefined();
    // Body text on a vertical page keeps the rotated overlay placement.
    expect(afterExact!.transform).toBe('rotate(90deg)');
    // The paragraph after the exact table starts one TABLE-WIDTH (50, the
    // physical width) past the flow top — NOT one logical-row-height (80) past:
    // place.left = cssW − (flowTop + tableW) = 200 − 80 = 120.
    expect(afterExact!.x).toBeCloseTo(CSS_W - (FLOW_TOP + TABLE_W), 1);
    // And the following auto table + paragraph keep flowing right→left.
    const afterAuto = runs.find((r) => r.text === 'qq');
    expect(afterAuto).toBeDefined();
    expect(afterAuto!.transform).toBe('rotate(90deg)');
    expect(afterAuto!.x).toBeLessThan(afterExact!.x);
  });

  it('pagination charges the physical table width as the flow footprint', () => {
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    // Logical flow budget = physical pageWidth − logical top/bottom insets
    // (30 + 24) = 146. A leading ~12 pt line + a 140 pt-EXACT-row table:
    // charging the logical row height (140) overflows onto page 2, while the
    // upright footprint (tableW = 50) fits page 1 beside it.
    const body = [
      { type: 'paragraph', ...bodyParagraph('zz') },
      { type: 'table', ...fixedTable(CELL_TEXT_1, 140) },
    ] as unknown as BodyElement[];
    const logicalSec = __test_verticalLayoutSection(PHYS);
    const pages = computePages(body, logicalSec, ctx, { 'Times New Roman': 'roman' });
    expect(pages.length).toBe(1);
    expect(pages[0].length).toBe(2);
  });

  it('retains upright physical geometry while the placement owns the vertical flow footprint', async () => {
    const doc = verticalTableDoc();
    const measure = makeRecordingCanvas();
    const logicalSec = __test_verticalLayoutSection(PHYS);
    const pages = computePages(
      doc.body,
      logicalSec,
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const tables = pages[0].filter((element) => element.type === 'table');
    expect(tables).toHaveLength(2);

    const exact = bodyFragmentFor(tables[0]);
    const auto = bodyFragmentFor(tables[1]);
    if (
      exact?.fragment.kind !== 'table' || !('flowBounds' in exact.fragment) ||
      auto?.fragment.kind !== 'table' || !('flowBounds' in auto.fragment)
    ) {
      throw new Error('expected retained upright table layouts');
    }

    // §17.6.20: the upright table's rows remain in physical coordinates, so
    // the exact row owns an 80 pt physical row stack. The surrounding vertical
    // story advances by the table's 50 pt physical width instead.
    expect(exact.fragment.advancePt).toBeCloseTo(80, 6);
    expect(auto.fragment.advancePt).toBeGreaterThan(80);
    expect(exact.heightPt).toBeCloseTo(TABLE_W, 6);
    expect(auto.heightPt).toBeCloseTo(TABLE_W, 6);

    const paint = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYS.pageWidth,
      prebuiltPages: pages,
    });
    expect(paint.measureCalls()).toBe(0);
  });

  it('records the destination column on an upright retained placement', () => {
    const wide = fixedTable('', 20);
    wide.colWidths = [140];
    wide.rows[0].cells[0].widthPt = 140;
    const narrow = fixedTable('', 20);
    const physical = {
      ...PHYS,
      columns: { count: 2, spacePt: 10, equalWidth: true, sep: false, cols: [] },
    } as SectionProps;
    const body = [
      { type: 'table', ...wide },
      { type: 'table', ...narrow },
    ] as unknown as BodyElement[];
    const measure = makeRecordingCanvas();
    const pages = computePages(
      body,
      __test_verticalLayoutSection(physical),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      { 'Times New Roman': 'roman' },
    );
    const tables = pages[0].filter((element) => element.type === 'table');
    expect(tables).toHaveLength(2);
    expect(bodyFragmentFor(tables[0])?.columnIndex).toBe(0);
    expect(bodyFragmentFor(tables[1])?.columnIndex).toBe(1);
  });
});
