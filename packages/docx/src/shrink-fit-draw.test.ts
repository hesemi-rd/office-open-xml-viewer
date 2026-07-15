import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import { SPACE_SHRINK_RATIO } from './line-layout.js';
import type {
  ShapeRun, ShapeText, ShapeTextRun,
  BodyElement, DocParagraph, DocxTextRun, DocxDocumentModel, SectionProps,
} from './types';

// Knuth-Plass shrink-fit: draw the line with the compression the fit judgment
// assumed. layoutLines admits a word onto the current line when the line's
// overflow Δ ≤ SPACE_SHRINK_RATIO · Σ(trailing-space widths) — i.e. it PROMISES
// the draw pass will squeeze the inter-word spaces by up to that budget. Before
// this fix the draw pass advanced the pen by the NATURAL widths, so an admitted
// line overran its box's clip and the last glyph was cut (sample-10 p1's centred
// text-box title "…Conference" lost its final "e"). These tests pin that a
// shrink-fit non-justified line (center / left / right; body AND text-box) is now
// squeezed back inside the box, a line that fits naturally is untouched, a
// justified line is not double-compressed, and the squeeze holds at a non-unit
// zoom (the #689 rescale path).
//
// Mock metric model (shared with the other draw-path tests): measureText width =
// code-point count × the font px, so every glyph — including a space — is exactly
// `px` wide. That makes the fit budget exact: one inter-word space = px, so
// Σtrailing = nSpaces·px and the admitted overflow is ≤ SPACE_SHRINK_RATIO·nSpaces·px.

const FONT_PX = 10;

interface FillTextEvent { text: string; x: number; y: number }

function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; fillTexts: FillTextEvent[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const ls = () => parseFloat(/(-?\d+(?:\.\d+)?)px/.exec(letterSpacing)?.[1] ?? '0');
  const fillTexts: FillTextEvent[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      const n = [...s].length;
      return {
        width: n * p + n * ls(),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { fillTexts.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { fillTexts.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fillTexts };
}

// ── Text-box (shape) path ────────────────────────────────────────────────────

const SHAPE_X = 100, SHAPE_Y = 50;

function shapeWith(blocks: ShapeText[]): ShapeRun {
  return {
    type: 'shape',
    presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: blocks,
  } as unknown as ShapeRun;
}

function block(text: string, extra: Partial<ShapeText> = {}): ShapeText {
  return {
    text, fontSizePt: FONT_PX, fontFamily: 'serif', alignment: 'left',
    runs: [{ text, fontSizePt: FONT_PX, fontFamily: 'serif' } as ShapeTextRun],
    ...extra,
  } as unknown as ShapeText;
}

function renderShape(blocks: ShapeText[], w: number, h: number, scale = 1): FillTextEvent[] {
  const { ctx, fillTexts } = makeRecordingCanvas();
  acquireAndPaintShapeTextBox(shapeWith(blocks), SHAPE_X * scale, SHAPE_Y * scale, w * scale, h * scale, ctx, scale);
  return fillTexts;
}

/** All fills on the single (top) line, left→right. */
function topLine(evs: FillTextEvent[]): FillTextEvent[] {
  const minY = Math.min(...evs.map((e) => e.y));
  return evs.filter((e) => e.y === minY).sort((a, b) => a.x - b.x);
}

/** Right edge of the last glyph on a line = its x + its trimmed measured width. */
function lineRight(line: FillTextEvent[]): number {
  const last = line[line.length - 1];
  const trimmed = last.text.replace(/ +$/, '');
  return last.x + [...trimmed].length * FONT_PX;
}

describe('shrink-fit draw — text-box (shape) path', () => {
  // "AAAA BBBB CCCC" = 5+5+4 = 140px natural; two spaces ⇒ Σtrailing = 20,
  // budget = 5px. Inner width 138 ⇒ overflow Δ = 2 ≤ 5, so layoutLines keeps all
  // three words on ONE line. The draw must squeeze the two spaces by 1px each so
  // the line's right edge lands on the box edge (138) instead of overrunning to
  // 140 and clipping the final "C".
  const CONTENT = 'AAAA BBBB CCCC';
  const NAT = 140; // 14 glyphs × 10
  const W = 138;

  it('squeezes a centred shrink-fit line so it lands inside the box (last glyph not clipped)', () => {
    const line = topLine(renderShape([block(CONTENT, { alignment: 'center' })], W, 400));
    // One line (all three words admitted).
    expect(line.map((e) => e.text.replace(/ +$/, '')).join(' ')).toBe('AAAA BBBB CCCC');
    const right = lineRight(line);
    // Lands on the inner-right edge, NOT at the natural 140 that would overrun the
    // box's clip. INNER_X = SHAPE_X (no insets).
    expect(right).toBeLessThanOrEqual(SHAPE_X + W + 0.01);
    expect(right).toBeCloseTo(SHAPE_X + W, 1);
    // The total squeeze equals the overflow (Δ = NAT − W = 2px), within the budget.
    const squeeze = NAT - (right - line[0].x);
    expect(squeeze).toBeCloseTo(NAT - W, 1);
    expect(squeeze).toBeLessThanOrEqual(SPACE_SHRINK_RATIO * 20 + 0.01);
  });

  it('squeezes a left-aligned shrink-fit line the same way (starts at inner-left, ends on edge)', () => {
    const line = topLine(renderShape([block(CONTENT, { alignment: 'left' })], W, 400));
    expect(line[0].x).toBeCloseTo(SHAPE_X, 3); // left edge unmoved
    expect(lineRight(line)).toBeLessThanOrEqual(SHAPE_X + W + 0.01);
    expect(lineRight(line)).toBeCloseTo(SHAPE_X + W, 1);
  });

  it('squeezes a right-aligned shrink-fit line to end exactly on the inner-right edge', () => {
    const line = topLine(renderShape([block(CONTENT, { alignment: 'right' })], W, 400));
    expect(lineRight(line)).toBeCloseTo(SHAPE_X + W, 1);
  });

  it('does NOT compress a line that already fits (spaces keep their natural width)', () => {
    // Same content, inner width 200 > natural 140 ⇒ positive slack, no squeeze.
    const line = topLine(renderShape([block(CONTENT, { alignment: 'center' })], 200, 400));
    // Inter-word gaps stay at exactly one space (10px): "BBBB" starts one space
    // past the end of "AAAA" glyphs.
    const a = line[0], b = line[1];
    const gap = b.x - (a.x + 4 * FONT_PX); // AAAA is 4 glyphs
    expect(gap).toBeCloseTo(FONT_PX, 3); // full, uncompressed space
    // And the whole line is NATURAL width, merely centred (right − left = 140).
    expect(lineRight(line) - line[0].x).toBeCloseTo(NAT, 1);
  });

  it('does NOT double-compress a justified (both) line — its own §17.18.44 path owns the slack', () => {
    // ECMA-376 §17.18.44 (issue #698): a `both` line breaks at its NATURAL fit —
    // the space-shrink tolerance is suppressed for justified paragraphs, so a
    // justified line never overflows its box and is never admitted with negative
    // slack. Line 1 is therefore a justify candidate with POSITIVE slack that the
    // justify path alone EXPANDS evenly; the shrink-fit branch (gated on
    // !applyJustify) stays off, so the gaps are single-application, not squeezed.
    // "AAAA BBBB CCCC DDDD EEEE" at WJ=160 breaks with line 1 = "AAAA BBBB CCCC "
    // (natural 140 ≤ 160), "DDDD" wrapping (140+40 > 160).
    const WJ = 160;
    const evs = renderShape([block('AAAA BBBB CCCC DDDD EEEE', { alignment: 'both' })], WJ, 400);
    const ys = [...new Set(evs.map((e) => e.y))].sort((a, b) => a - b);
    expect(ys.length).toBeGreaterThanOrEqual(2); // it wrapped
    const line1 = evs.filter((e) => e.y === ys[0]).sort((a, b) => a.x - b.x);
    expect(line1.map((e) => e.text.replace(/ +$/, '')).join(' ')).toBe('AAAA BBBB CCCC');
    // Justify signature: the inter-word gaps are EXPANDED by an EQUAL amount
    // (Σ slack distributed evenly), so the two realised gaps match. If the shrink-
    // fit path ALSO fired it would compress the gaps below their natural width or
    // distribute unevenly — so equal gaps WIDER than one space prove that only the
    // justify pass ran.
    const [a, b, c] = line1; // "AAAA ", "BBBB ", "CCCC "
    const gap1 = b.x - (a.x + 4 * FONT_PX); // realised inter-word gap 1
    const gap2 = c.x - (b.x + 4 * FONT_PX); // realised inter-word gap 2
    expect(gap1).toBeCloseTo(gap2, 3);      // even distribution (justify)
    // Expanded from the natural 10px by a SINGLE justify pass (positive slack),
    // never squeezed by the shrink-fit branch.
    expect(gap1).toBeGreaterThan(FONT_PX);
  });

  it('holds the squeeze at a non-unit zoom (scale = 0.75) — line stays inside the scaled box', () => {
    const scale = 0.75;
    const line = topLine(renderShape([block(CONTENT, { alignment: 'center' })], W, 400, scale));
    // Box inner-right at this scale.
    const innerRight = (SHAPE_X + W) * scale;
    expect(lineRight2(line, scale)).toBeLessThanOrEqual(innerRight + 0.5);
    expect(lineRight2(line, scale)).toBeCloseTo(innerRight, 0);
  });
});

/** lineRight with an explicit scale (glyph advance = px·scale under the mock). */
function lineRight2(line: FillTextEvent[], scale: number): number {
  const last = line[line.length - 1];
  const trimmed = last.text.replace(/ +$/, '');
  return last.x + [...trimmed].length * FONT_PX * scale;
}

// ── Body path ────────────────────────────────────────────────────────────────

function bodyTextRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
  } as DocxTextRun;
}

function bodyPara(text: string, alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...bodyTextRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics', widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function bodySection(pageWidth: number): SectionProps {
  return {
    pageWidth, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

async function renderBody(body: BodyElement[], sec: SectionProps): Promise<FillTextEvent[]> {
  const { ctx, fillTexts } = makeRecordingCanvas();
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx } as unknown as HTMLCanvasElement;
  const model = {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
  await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: sec.pageWidth });
  return fillTexts;
}

describe('shrink-fit draw — body paragraph path', () => {
  // "AAAA BBBB CCCC DDDD" = 5+5+5+4 = 190px natural; three spaces ⇒ Σtrailing = 30,
  // budget = 7.5px. Page width 185 (margins 0) ⇒ overflow Δ = 5 ≤ 7.5 ⇒ one line.
  const CONTENT = 'AAAA BBBB CCCC DDDD';
  const NAT = 190;
  const PAGE = 185;

  it('squeezes a centred shrink-fit body line inside the text margin (last glyph not clipped)', async () => {
    const evs = await renderBody([bodyPara(CONTENT, 'center')], bodySection(PAGE));
    const line = topLine(evs);
    expect(line.map((e) => e.text.replace(/ +$/, '')).join(' ')).toBe('AAAA BBBB CCCC DDDD');
    const right = lineRight(line);
    // Margin left = 0, avail = PAGE. The last glyph must land within the margin.
    expect(right).toBeLessThanOrEqual(PAGE + 0.01);
    expect(right).toBeCloseTo(PAGE, 1);
    // Squeeze == overflow, within budget.
    expect(NAT - (right - line[0].x)).toBeCloseTo(NAT - PAGE, 1);
  });

  it('does NOT compress a body line that already fits', async () => {
    const evs = await renderBody([bodyPara(CONTENT, 'center')], bodySection(260));
    const line = topLine(evs);
    // Natural width preserved (merely centred).
    expect(lineRight(line) - line[0].x).toBeCloseTo(NAT, 1);
  });
});
