import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.2.35 `<w:spacing>` / §17.3.2.43 `<w:w>` must reach the actual
// GLYPH DRAW on every paint branch, and the FOLLOWING run's pen position must
// coincide with the previous run's painted extent (measure==paint at the run
// boundary, the sample-34 overlap contract). The advance-model unit tests
// (charspacing-advance.test.ts) derive both sides from segAdvanceWidth, so they
// cannot catch a paint branch that forgets ctx.letterSpacing or advances glyphs
// by the natural width — these tests drive renderDocumentToCanvas end-to-end
// with a recording canvas and assert on the recorded fillText calls instead.
//
// Four paint branches (renderer.ts drawLine segment loop):
//   1. horizontal common single-fillText (no grid, no justify)
//   2. horizontal docGrid character-grid branch (grid pitch + char spacing)
//   3. horizontal §17.18.44 justify branch (distributed pitch + char spacing)
//   4. vertical tbRl branch (drawVerticalRun per-glyph advance)
// ─────────────────────────────────────────────────────────────────────────────

const FONT_PX = 20;

interface FillCall {
  text: string;
  x: number;
  y: number;
  letterSpacing: string;
  // Naive accumulation of the raw translate() x arguments in effect at fill
  // time (save/restore-scoped, rotation ignored). Constant page-level offsets
  // cancel in the DELTAS between fills of the same line, which is all the
  // vertical assertions read (drawVerticalRun brackets each upright glyph in
  // save; translate(cellCenter, baseline); rotate; fillText; restore).
  translateX: number;
  scaleX: number;
  scaleY: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: FillCall[] = [];
  let scaleX = 1;
  let scaleY = 1;
  let translateX = 0;
  const stack: { scaleX: number; scaleY: number; translateX: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    fontKerning: 'auto',
    measureText: (s: string) => {
      const p = px();
      const w = [...s].length * p;
      return {
        width: w,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ scaleX, scaleY, translateX }); },
    restore() { const s = stack.pop(); if (s) { scaleX = s.scaleX; scaleY = s.scaleY; translateX = s.translateX; } },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    scale(sx: number, sy?: number) { scaleX *= sx; scaleY *= sy ?? sx; },
    translate(tx: number) { translateX += tx; },
    rotate() {},
    fillText(text: string, x: number, y: number) {
      fills.push({ text, x, y, letterSpacing, translateX, scaleX, scaleY });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocxTextRun[], alignment = 'left'): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(extra: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 600, pageHeight: 600, marginTop: 0, marginRight: 0, marginBottom: 0,
    marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
    evenAndOddHeaders: false, ...extra,
  } as SectionProps;
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(
  body: BodyElement[],
  sec: SectionProps = section(),
): Promise<FillCall[]> {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, { dpr: 1, width: 600 });
  return fills;
}

function drawOf(fills: FillCall[], text: string): FillCall {
  const f = fills.find((c) => c.text === text);
  expect(f, `a fillText drew ${JSON.stringify(text)}`).toBeDefined();
  return f as FillCall;
}

/** Effective x (page space) of the fills the vertical per-glyph draw produced:
 *  the accumulated translate at fill time. Constant page-level offsets cancel
 *  when the caller diffs consecutive values. */
function glyphCenters(fills: FillCall[], chars: string[]): number[] {
  return chars.map((ch) => drawOf(fills, ch).translateX);
}

describe('run charSpacing/charScale reach the painted glyphs on every branch', () => {
  it('horizontal common path: following run starts at the expanded right edge', async () => {
    // run1 4 cps × (20 natural + 2 spacing) = 88; run2 must start +88 from run1.
    const fills = await render([para([
      textRun('あいうえ', { charSpacing: 2 }),
      textRun('かき'),
    ])]);
    const r1 = drawOf(fills, 'あいうえ');
    const r2 = drawOf(fills, 'かき');
    expect(r1.letterSpacing).toBe('2px'); // §17.3.2.35 at the glyph draw
    expect(r2.letterSpacing).toBe('0px');
    expect(r2.x - r1.x).toBeCloseTo(4 * FONT_PX + 4 * 2, 5);
  });

  it('docGrid character-grid path: grid pitch and charSpacing combine at the draw', async () => {
    // §17.6.5 linesAndChars with charSpace 4096 units = 1 pt per EA glyph.
    // Combined per-glyph pitch = 1 (grid) + 2 (w:spacing) = 3px on run1; the
    // pure-EA run2 keeps its own 1px grid pitch. run2 starts at run1's celled
    // advance: 4 × 20 + 4 × 3 = 92.
    const fills = await render(
      [para([textRun('あいうえ', { charSpacing: 2 }), textRun('かき')])],
      section({ docGridType: 'linesAndChars', docGridLinePitch: 24, docGridCharSpace: 4096 }),
    );
    const r1 = drawOf(fills, 'あいうえ');
    const r2 = drawOf(fills, 'かき');
    expect(r1.letterSpacing).toBe('3px');
    expect(r2.letterSpacing).toBe('1px');
    expect(r2.x + r2.translateX - (r1.x + r1.translateX)).toBeCloseTo(4 * FONT_PX + 4 * 3, 5);
  });

  it('justify path: charSpacing adds to the distributed pitch at the draw', async () => {
    // A justified ('both') CJK paragraph that wraps: line 1 carries run1 (25 cps,
    // w:spacing 2) + a 2-cp prefix of run2, with 10px of slack distributed over
    // the inter-CJK gaps. Acquisition retains contextual text and a uniform
    // pitch. The two runs' recorded pitch differs by exactly the authored 2px.
    const fills = await render([para(
      [textRun('あ'.repeat(25), { charSpacing: 2 }), textRun('い'.repeat(10))],
      'both',
    )]);
    const r1 = fills.find((f) => f.text.startsWith('あ'));
    const r2 = fills.find((f) => f.text.startsWith('い'));
    expect(r1, 'run1 glyphs painted').toBeDefined();
    expect(r2, 'run2 first-line glyphs painted').toBeDefined();
    const p1 = parseFloat((r1 as FillCall).letterSpacing);
    const p2 = parseFloat((r2 as FillCall).letterSpacing);
    expect((r1 as FillCall).text).toBe('あ'.repeat(25));
    expect(p1 - p2).toBeCloseTo(2, 5); // the §17.3.2.35 component survives justify
  });

  it('vertical tbRl path: per-glyph advance includes charSpacing (no gap before the next run)', async () => {
    // ECMA-376 §17.6.20 rotate-layout: the logical line axis is the along-column
    // advance. layoutLines gives run1 measuredWidth = 4 × 20 + 4 × 2 = 88, so the
    // pen places run2 at 88. drawVerticalRun must advance each upright glyph by
    // measure + 2 (the §17.3.2.35 pitch); with the natural-only advance run1's
    // glyphs end at 80 and an 8px hole opens before run2.
    const fills = await render(
      [para([textRun('あいうえ', { charSpacing: 2 }), textRun('かき')])],
      section({ textDirection: 'tbRl' }),
    );
    const [a, i, u, e, ka] = glyphCenters(fills, ['あ', 'い', 'う', 'え', 'か']);
    // Cell centres of run1's glyphs are (20+2) apart.
    expect(i - a).toBeCloseTo(22, 5);
    expect(u - i).toBeCloseTo(22, 5);
    expect(e - u).toBeCloseTo(22, 5);
    // Run boundary: run2's first cell centre sits half-cells from run1's last —
    // (22 + 20) / 2 = 21 — i.e. flush, no hole (pen 88 == painted extent 88).
    expect(ka - e).toBeCloseTo((22 + 20) / 2, 5);
  });

  it('vertical tbRl path: w:w scales the along-column advance (measure==paint)', async () => {
    // §17.3.2.43 w:w in the rotate-layout vertical engine: the layout kernel is
    // direction-agnostic and has always scaled the LINE-AXIS advance by w:w
    // (segAdvanceWidth), which after the +90° page rotation IS the along-column
    // cell extent — so wrap points, run boxes, and find/selection already assume
    // it. Paint must follow measure: run1 (w:w=50%) cells are 10px, and run2's
    // pen (= run1 measuredWidth = 40) sits flush against the last cell.
    const fills = await render(
      [para([textRun('あいうえ', { charScale: 0.5 }), textRun('かき')])],
      section({ textDirection: 'tbRl' }),
    );
    const [a, i, u, e, ka] = glyphCenters(fills, ['あ', 'い', 'う', 'え', 'か']);
    expect(i - a).toBeCloseTo(10, 5);
    expect(u - i).toBeCloseTo(10, 5);
    expect(e - u).toBeCloseTo(10, 5);
    // Boundary: (10 + 20) / 2 = 15. With the unscaled per-glyph advance run1's
    // glyphs overrun their cells and run2 paints INSIDE run1's tail.
    expect(ka - e).toBeCloseTo((10 + 20) / 2, 5);
  });
});
