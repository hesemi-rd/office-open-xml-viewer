import { describe, it, expect } from 'vitest';
import {
  renderDocumentToCanvas,
  __test_setTableReuseEnabled,
  __test_setFragmentPaintEnabled,
} from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Finding 1 — cell height measurement must use REAL-SCALE Canvas metrics.
//
// The cell-height path (measureCellParagraphHeight → measureCellElementHeight →
// measureCellContentHeightPx) measures a cell paragraph at scale 1 and then does
// a geometric `× scale` on the scale-1 point height. That is the exact anti-
// pattern `rescaleLayoutLines` exists to avoid: a real (hinted) font's Canvas
// metrics are NOT scale-linear, so `metric(pt · s) ≠ s · metric(pt)`. The paint
// path (renderParagraph) rehydrates the scale-1 line PARTITION to the paint scale
// by RE-MEASURING every line at that scale (rescaleLayoutLines), so the painted
// content occupies a height the naive `× scale` cannot reproduce. When the two
// disagree, a vAlign=center/bottom cell centres its content off the true middle,
// and the content-driven row-height fallback reserves the wrong height.
//
// These tests mock a NON-LINEAR `measureText` — a sub-linear ascent/descent, the
// direction real font hinting bends (glyphs proportionally shorter at larger
// sizes) — so `metric(12·s) ≠ s·metric(12)`. They then render at scale = 4/3
// (≈1.333, the renderer's default cssWidth/physWidth ratio) and assert the cell
// content height that vAlign / row layout USES equals the height the paint side
// actually PRODUCES, read back through the public render seam.
// ─────────────────────────────────────────────────────────────────────────────

/** Sub-linear single-glyph vertical metrics in px at font size `p` px. The
 *  ascent/descent shrink proportionally as `p` grows, so `metric(p·s) ≠ s·metric(p)`
 *  — the hinting-like non-linearity `rescaleLayoutLines` re-measures for. Width is
 *  kept linear (charCount · p · 0.5) so the line PARTITION never changes with
 *  scale; only the per-line BOX height is scale-non-linear, which is exactly the
 *  quantity that drives a cell's measured height. */
function glyphMetrics(p: number): { asc: number; desc: number } {
  return { asc: p * (0.8 - 0.01 * p), desc: p * (0.2 - 0.005 * p) };
}

interface FillTextCall { text: string; x: number; y: number; font: string; }
interface FillRectCall { x: number; y: number; w: number; h: number; fillStyle: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: FillTextCall[];
  fillRectCalls: FillRectCall[];
} {
  let font = '10px serif';
  let fillStyle = '#000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fillTextCalls: FillTextCall[] = [];
  const fillRectCalls: FillRectCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      const { asc, desc } = glyphMetrics(p);
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: asc,
        fontBoundingBoxDescent: desc,
        actualBoundingBoxAscent: asc,
        actualBoundingBoxDescent: desc,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect(x: number, y: number, w: number, h: number) {
      fillRectCalls.push({ x, y, w, h, fillStyle });
    },
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, font });
    },
    strokeText() {},
    strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  // `clipExact` rows (ST_HeightRule exact) read `ctx.canvas.width`; wire the
  // back-reference so the exact-height cells below render.
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls, fillRectCalls };
}

// An untabled synthetic font so the font-metrics single-line FLOOR
// (intendedSingleLinePx) is 0 and the line box is exactly the mock's ascent +
// descent — isolating the scale non-linearity under test.
const TEST_FONT = 'Synthetic Untabled Serif';

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  };
}

function paraOf(text: string, opts: Partial<DocParagraph> = {}): CellElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 12, defaultFontFamily: TEST_FONT,
    widowControl: false,
    ...opts,
  } as unknown as CellElement;
}

function cell(
  content: CellElement[],
  vAlign: 'top' | 'center' | 'bottom',
  background: string | null = null,
): DocTableCell {
  return {
    content,
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background,
    vAlign,
    widthPt: 300,
  } as DocTableCell;
}

function row(c: DocTableCell, rowHeight: number | null, rule: 'auto' | 'atLeast' | 'exact'): DocTableRow {
  return {
    cells: [c],
    rowHeight,
    rowHeightRule: rule,
    isHeader: false,
  } as DocTableRow;
}

function tableOf(r: DocTableRow): DocTable {
  return {
    colWidths: [300],
    rows: [r],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  } as DocTable;
}

function docWithTable(t: DocTable): DocxDocumentModel {
  return {
    section: {
      pageWidth: 300, pageHeight: 600,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...t } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

// scale = cssWidth / pageWidthPt = 400 / 300 = 4/3 ≈ 1.333.
const CSS_WIDTH = 400;
const PAGE_WIDTH = 300;
const SCALE = CSS_WIDTH / PAGE_WIDTH;
const FONT_PX = 12 * SCALE; // 16 exactly

async function renderAndRead(t: DocTable) {
  const rec = makeRecordingCanvas();
  await renderDocumentToCanvas(docWithTable(t), rec.canvas, 0, {
    dpr: 1,
    width: CSS_WIDTH,
  });
  return rec;
}

/** Painted inked-block extent (px) of a set of one-line paragraphs, read from the
 *  fillText baselines and the mock's paint-scale glyph box. This is the height the
 *  paint side actually PRODUCES — the reference every measured height is compared
 *  against. */
function paintedExtent(calls: FillTextCall[], firstText: string, lastText: string): { top: number; bottom: number } {
  const first = calls.find((c) => c.text === firstText);
  const last = calls.find((c) => c.text === lastText);
  expect(first).toBeDefined();
  expect(last).toBeDefined();
  const { asc, desc } = glyphMetrics(FONT_PX);
  return { top: first!.y - asc, bottom: last!.y + desc };
}

describe('Finding 1 — cell content height uses real-scale (rescaled) Canvas metrics', () => {
  it('sanity: the mock ascent/descent is scale-NON-linear (else the test is vacuous)', () => {
    const one = glyphMetrics(12);
    const s = glyphMetrics(12 * SCALE);
    const oneH = one.asc + one.desc;
    const scaledH = s.asc + s.desc;
    // A geometric ×scale of the scale-1 box must MISS the real paint-scale box by
    // a clearly observable margin — this is the divergence Finding 1 is about.
    expect(Math.abs(oneH * SCALE - scaledH)).toBeGreaterThan(0.5);
  });

  it('vAlign=center: the inked block is centred using the PAINTED content height', async () => {
    // Three single-line paragraphs, no spacing. Exact 120 pt row → drawn height is
    // exactly 120·scale px (resolveSingleRowHeight, ST_HeightRule exact) and is
    // NOT content-measured, so this isolates the vAlign centring consumer. The
    // content (~3 lines) fits well within the row, leaving centring room.
    const t = tableOf(row(
      cell(
        [paraOf('aa'), paraOf('bb'), paraOf('cc')],
        'center',
      ),
      120,
      'exact',
    ));
    const { fillTextCalls } = await renderAndRead(t);
    const { top, bottom } = paintedExtent(fillTextCalls, 'aa', 'cc');
    const inkedMid = (top + bottom) / 2;
    const cellMid = (120 * SCALE) / 2; // row at y=0, exact height 120·scale.
    // measure == paint: the vAlign offset used the true painted content height, so
    // the painted block's midpoint lands on the cell midpoint. With the geometric
    // ×scale bug the block is off-centre by ~1.4 px here.
    expect(inkedMid).toBeCloseTo(cellMid, 1);
  });

  it('vAlign=bottom: the inked block bottom hugs the cell bottom (painted height)', async () => {
    const t = tableOf(row(
      cell([paraOf('aa'), paraOf('bb'), paraOf('cc')], 'bottom'),
      120,
      'exact',
    ));
    const { fillTextCalls } = await renderAndRead(t);
    const { bottom } = paintedExtent(fillTextCalls, 'aa', 'cc');
    // mb = 0 → the inked bottom sits on the cell bottom (120·scale). A geometric
    // ×scale measured height would seat it short of / past the true bottom.
    expect(bottom).toBeCloseTo(120 * SCALE, 1);
  });

  it('auto row height fallback: reserved row height equals the painted content height', async () => {
    // Auto height → the row height IS the measured cell content height. Disable the
    // stamped-layout reuse so computeTableLayout runs the fresh real-scale fallback
    // (the exact path Finding 1 flags), not the paginator's scale-1 stamp × scale.
    // PR 6 — also disable the fragment paint path (which, like the reuse, draws the
    // paginator's scale-1 row height × scale); the legacy `renderTable` recompute is
    // the path this Finding characterizes.
    const prevFrag = __test_setFragmentPaintEnabled(false);
    const prev = __test_setTableReuseEnabled(false);
    try {
      const t = tableOf(row(
        cell([paraOf('aa'), paraOf('bb'), paraOf('cc')], 'top', 'abcdef'),
        null,
        'auto',
      ));
      const { fillTextCalls, fillRectCalls } = await renderAndRead(t);
      const bg = fillRectCalls.find((r) => r.fillStyle === '#abcdef');
      expect(bg).toBeDefined();
      const { top, bottom } = paintedExtent(fillTextCalls, 'aa', 'cc');
      // Zero cell margins → the reserved (drawn) row height must equal the painted
      // content extent. The geometric ×scale fallback reserves the scale-1 height
      // × scale, which misses the painted height under non-linear metrics.
      expect(bg!.h).toBeCloseTo(bottom - top, 1);
    } finally {
      __test_setTableReuseEnabled(prev);
      __test_setFragmentPaintEnabled(prevFrag);
    }
  });
});
