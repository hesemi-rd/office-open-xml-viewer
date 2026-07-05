import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  LineNumbering,
  PageBorders,
  SectionProps,
} from './types';

// ECMA-376 §17.6.10 (pgBorders), §17.6.8 (lnNumType), §17.6.23 (vAlign) — the
// three sectPr page decorations. These tests render a synthetic single-section
// document through `renderDocumentToCanvas` and read the recording canvas to
// pin the DRAWN geometry (border rectangle, line-number glyphs, body offset).

const TEST_FONT = 'Synthetic Untabled Serif';

interface FillTextCall { text: string; x: number; y: number; font: string; align: string; }
interface Seg { x1: number; y1: number; x2: number; y2: number; }

// A crisp-snapped border stroke may be nudged by up to ~0.5 px perpendicular to
// its direction (see strokeCrispSegment / crispOffset). Match a horizontal rule
// whose y is within 1 px of the target, and likewise a vertical rule's x.
function hasHoriz(segs: Seg[], y: number): boolean {
  return segs.some((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.y1 - y) <= 1);
}
function hasVert(segs: Seg[], x: number): boolean {
  return segs.some((s) => Math.abs(s.x1 - s.x2) < 0.5 && Math.abs(s.x1 - x) <= 1);
}

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: FillTextCall[];
  segments: Seg[];
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fillTextCalls: FillTextCall[] = [];
  const segments: Seg[] = [];
  let cur = { x: 0, y: 0 };
  let align: CanvasTextAlign = 'left';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get textAlign() { return align; },
    set textAlign(v: CanvasTextAlign) { align = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
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
    moveTo(x: number, y: number) { cur = { x, y }; },
    lineTo(x: number, y: number) { segments.push({ x1: cur.x, y1: cur.y, x2: x, y2: y }); cur = { x, y }; },
    stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, font, align });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls, segments };
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
}

function para(text: string, opts: Partial<DocParagraph> = {}): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) }],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT,
    widowControl: false,
    ...opts,
  } as unknown as BodyElement;
}

function section(over: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 40,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...over,
  } as SectionProps;
}

function docOf(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(doc: DocxDocumentModel, pageIndex = 0) {
  const rec = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, rec.canvas, pageIndex, {
    dpr: 1,
    width: 200, // scale = 1 (px per pt) ⇒ asserts are in pt units
  });
  return rec;
}

// ============================================================================
// §17.6.10 — page borders
// ============================================================================

describe('§17.6.10 pgBorders — page border rectangle', () => {
  const edge = (over = {}) => ({ style: 'single', color: 'ff0000', width: 1, space: 24, ...over });

  it('offsetFrom="page": each edge is inset from the PAGE edge by its space (pt)', async () => {
    const pb: PageBorders = {
      offsetFrom: 'page', display: 'allPages', zOrder: 'front',
      top: edge(), bottom: edge(), left: edge(), right: edge(),
    };
    const { segments } = await render(docOf([para('body')], section({ pageBorders: pb })));
    // Four axis-aligned segments forming a rectangle inset 24 pt from every page
    // edge. Page is 200×200 at scale 1 ⇒ rectangle [24,24]–[176,176].
    expect(hasHoriz(segments, 24)).toBe(true);   // top
    expect(hasHoriz(segments, 176)).toBe(true);  // bottom (200 - 24)
    expect(hasVert(segments, 24)).toBe(true);    // left
    expect(hasVert(segments, 176)).toBe(true);   // right
  });

  it('offsetFrom="text": each edge is inset from the TEXT MARGIN by its space', async () => {
    const pb: PageBorders = {
      offsetFrom: 'text', display: 'allPages', zOrder: 'front',
      top: edge({ space: 4 }), left: edge({ space: 4 }),
      bottom: edge({ space: 4 }), right: edge({ space: 4 }),
    };
    // margins: top/bottom 20, left 40, right 20 ⇒ text box [40,20]–[180,180].
    // +4 pt inset ⇒ rectangle [44,24]–[176,176].
    const { segments } = await render(docOf([para('body')], section({ pageBorders: pb })));
    expect(hasHoriz(segments, 24)).toBe(true);   // top margin 20 + 4
    expect(hasHoriz(segments, 176)).toBe(true);  // bottom margin (200-20) - 4
    expect(hasVert(segments, 44)).toBe(true);    // left margin 40 + 4
    expect(hasVert(segments, 176)).toBe(true);   // right margin (200-20) - 4
  });

  it('display="firstPage" shows the border only on page 0', async () => {
    const pb: PageBorders = { offsetFrom: 'page', display: 'firstPage', zOrder: 'front', top: edge() };
    const twoPages = [para('a'), { type: 'pageBreak' } as BodyElement, para('b')];
    const p0 = await render(docOf(twoPages, section({ pageBorders: pb })), 0);
    const p1 = await render(docOf(twoPages, section({ pageBorders: pb })), 1);
    expect(hasHoriz(p0.segments, 24)).toBe(true);
    expect(hasHoriz(p1.segments, 24)).toBe(false);
  });

  it('display="notFirstPage" hides the border on page 0, shows it later', async () => {
    const pb: PageBorders = { offsetFrom: 'page', display: 'notFirstPage', zOrder: 'front', top: edge() };
    const twoPages = [para('a'), { type: 'pageBreak' } as BodyElement, para('b')];
    const p0 = await render(docOf(twoPages, section({ pageBorders: pb })), 0);
    const p1 = await render(docOf(twoPages, section({ pageBorders: pb })), 1);
    expect(hasHoriz(p0.segments, 24)).toBe(false);
    expect(hasHoriz(p1.segments, 24)).toBe(true);
  });

  it('no pgBorders ⇒ no border rectangle drawn (non-regression)', async () => {
    const { segments } = await render(docOf([para('body')], section()));
    // A plain paragraph draws no long axis-aligned rules at a border inset.
    expect(hasHoriz(segments, 24)).toBe(false);
  });
});

// ============================================================================
// §17.6.8 — line numbering
// ============================================================================

describe('§17.6.8 lnNumType — line numbering', () => {
  const ln = (over: Partial<LineNumbering> = {}): LineNumbering =>
    ({ countBy: 1, start: 1, distance: 10, restart: 'newPage', ...over });

  it('countBy=1 numbers every body line in the left margin, right-aligned', async () => {
    const doc = docOf([para('one'), para('two'), para('three')], section({ lineNumbering: ln() }));
    const { fillTextCalls } = await render(doc);
    const numbers = fillTextCalls.filter((c) => /^\d+$/.test(c.text));
    expect(numbers.map((n) => n.text)).toEqual(['1', '2', '3']);
    // Right edge sits distance(10) pt left of the text margin (marginLeft 40) ⇒ x=30.
    for (const n of numbers) {
      expect(n.x).toBeCloseTo(30, 2);
      expect(n.align).toBe('right');
    }
  });

  it('countBy=2 numbers only every 2nd line', async () => {
    const doc = docOf(
      [para('a'), para('b'), para('c'), para('d')],
      section({ lineNumbering: ln({ countBy: 2 }) }),
    );
    const { fillTextCalls } = await render(doc);
    const numbers = fillTextCalls.filter((c) => /^\d+$/.test(c.text)).map((n) => n.text);
    expect(numbers).toEqual(['2', '4']);
  });

  it('start=5 begins the counter at 5', async () => {
    const doc = docOf([para('a'), para('b')], section({ lineNumbering: ln({ start: 5 }) }));
    const { fillTextCalls } = await render(doc);
    const numbers = fillTextCalls.filter((c) => /^\d+$/.test(c.text)).map((n) => n.text);
    expect(numbers).toEqual(['5', '6']);
  });

  it('restart="newPage" restarts the counter at `start` on every page', async () => {
    const body = [para('a'), para('b'), { type: 'pageBreak' } as BodyElement, para('c'), para('d')];
    const doc = docOf(body, section({ lineNumbering: ln({ restart: 'newPage' }) }));
    const p1 = await render(doc, 1);
    const numbers = p1.fillTextCalls.filter((c) => /^\d+$/.test(c.text)).map((n) => n.text);
    // Page 1 restarts at 1 (newPage), so its two lines are 1,2 (NOT 3,4).
    expect(numbers).toEqual(['1', '2']);
  });

  it('restart="continuous" continues the counter across pages', async () => {
    const body = [para('a'), para('b'), { type: 'pageBreak' } as BodyElement, para('c'), para('d')];
    const doc = docOf(body, section({ lineNumbering: ln({ restart: 'continuous' }) }));
    const p1 = await render(doc, 1);
    const numbers = p1.fillTextCalls.filter((c) => /^\d+$/.test(c.text)).map((n) => n.text);
    // Page 0 drew lines 1,2 ⇒ page 1 continues at 3,4.
    expect(numbers).toEqual(['3', '4']);
  });

  it('no lnNumType ⇒ no line numbers drawn (non-regression)', async () => {
    const { fillTextCalls } = await render(docOf([para('a'), para('b')], section()));
    expect(fillTextCalls.filter((c) => /^\d+$/.test(c.text))).toHaveLength(0);
  });
});

// ============================================================================
// §17.6.23 — vertical alignment
// ============================================================================

describe('§17.6.23 vAlign — body vertical alignment', () => {
  // One 10 pt line: ascent 8, descent 2. Body band = [20, 180] (margins 20/20)
  // ⇒ band height 160, content height ~10.
  it('vAlign="center" shifts the body to the vertical centre of the text band', async () => {
    const top = await render(docOf([para('x')], section()));
    const centered = await render(docOf([para('x')], section({ vAlign: 'center' })));
    const topX = top.fillTextCalls.find((c) => c.text === 'x')!;
    const midX = centered.fillTextCalls.find((c) => c.text === 'x')!;
    // Top-aligned baseline sits near the top margin (20 + ascent 8 = 28).
    expect(topX.y).toBeCloseTo(28, 1);
    // Centred baseline sits ~ (band centre 100) − halfLine + ascent ≈ 20 + 75 + 8.
    // Band 160, content 10 ⇒ shift (160-10)/2 = 75. New top 20+75=95, baseline 95+8=103.
    expect(midX.y).toBeCloseTo(103, 1);
  });

  it('vAlign="bottom" pushes the body to the bottom margin', async () => {
    const bottom = await render(docOf([para('x')], section({ vAlign: 'bottom' })));
    const bx = bottom.fillTextCalls.find((c) => c.text === 'x')!;
    // Shift = band(160) − content(10) = 150. New top 20+150=170, baseline 170+8=178.
    // Inked bottom = baseline + descent(2) = 180 = bottom margin.
    expect(bx.y + 10 * 0.2).toBeCloseTo(180, 1);
  });

  it('vAlign="top" (default) leaves the body at the top margin (non-regression)', async () => {
    const noAlign = await render(docOf([para('x')], section()));
    const topAlign = await render(docOf([para('x')], section({ vAlign: 'top' })));
    const a = noAlign.fillTextCalls.find((c) => c.text === 'x')!;
    const b = topAlign.fillTextCalls.find((c) => c.text === 'x')!;
    expect(a.y).toBeCloseTo(b.y, 2);
  });
});
