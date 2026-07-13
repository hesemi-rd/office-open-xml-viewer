import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { paragraphMarkLineHeight } from './line-layout.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.33 (line height) + §17.3.1.29 (paragraph mark): an EMPTY
// paragraph still occupies one paragraph-mark line, whose height is the mark
// font's single-line height — the SAME computation a text line of that font and
// size uses (layoutLines → fontBoundingBox). The renderer must not size the
// empty mark line from a synthetic 1-em box while text lines use the font's real
// (often substituted) metrics; that asymmetry under-measured every empty
// paragraph. In sample-12's figure block a run of nine empty "spacer" paragraphs
// fell ~1.65 pt short per line, so the following centered caption rose into the
// image float's wrap band and its first line wrapped beside the figure instead
// of clearing it (the whole caption sits below the figure in Word).
//
// This stub reports an ASYMMETRIC ~1.15-em font box (asc 0.95 + desc 0.20),
// mimicking a substituted Latin font (a real browser/skia "Times New Roman"
// fallback reports ~1.15 em, NOT the synthetic 1.0 em). With a 1.0-em stub the
// two code paths would coincide and the regression would be invisible.

interface FillTextCall { text: string; x: number; y: number; }

const ASC_RATIO = 0.95;
const DESC_RATIO = 0.2; // sum = 1.15 em ≠ the old synthetic 0.8 + 0.2 = 1.0 em

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: FillTextCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * ASC_RATIO,
        fontBoundingBoxDescent: p * DESC_RATIO,
        actualBoundingBoxAscent: p * ASC_RATIO,
        actualBoundingBoxDescent: p * DESC_RATIO,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 11, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocxTextRun;
}

function para(text: string): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function docOf(bodyParas: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 800,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: bodyParas as unknown as BodyElement[],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function baselines(bodyParas: DocParagraph[]): Promise<FillTextCall[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(docOf(bodyParas), canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('empty paragraph mark line height (§17.3.1.29 / §17.3.1.33)', () => {
  it('an empty paragraph advances the cursor by the SAME single-line height as a text paragraph', async () => {
    // Reference: [A, M, B] — three single-line text paragraphs. The gap between
    // A and B baselines = lineHeight(M) + lineHeight(A) (one full intervening line).
    const refCalls = await baselines([para('A'), para('M'), para('B')]);
    const refA = refCalls.find((c) => c.text === 'A')!;
    const refB = refCalls.find((c) => c.text === 'B')!;
    expect(refA).toBeDefined();
    expect(refB).toBeDefined();
    const refGap = refB.y - refA.y;

    // Subject: [A, <empty>, B] — the middle paragraph is empty. Its mark line must
    // reserve the same single-line height, so the A→B gap is identical.
    const subjCalls = await baselines([para('A'), para(''), para('B')]);
    const subjA = subjCalls.find((c) => c.text === 'A')!;
    const subjB = subjCalls.find((c) => c.text === 'B')!;
    expect(subjA).toBeDefined();
    expect(subjB).toBeDefined();
    const subjGap = subjB.y - subjA.y;

    // Before the fix subjGap was short by (1.15 − 1.0) × 11 = 1.65 pt because the
    // empty mark line used a synthetic 1-em box while the text lines used the
    // 1.15-em stub box.
    expect(subjGap).toBeCloseTo(refGap, 2);
  });

  it('uses the paragraph mark eastAsia font axis in East Asian document-grid context', () => {
    const p = {
      ...para(''),
      defaultFontFamily: 'Century',
      defaultFontFamilyEastAsia: 'ＭＳ 明朝',
    } as DocParagraph;
    let font = '';
    const ctx = {
      get font() { return font; },
      set font(v: string) { font = v; },
      measureText: () => {
        const ea = font.includes('ＭＳ 明朝');
        const total = ea ? 30 : 10;
        return {
          width: 0,
          fontBoundingBoxAscent: total * 0.8,
          fontBoundingBoxDescent: total * 0.2,
          actualBoundingBoxAscent: total * 0.8,
          actualBoundingBoxDescent: total * 0.2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(p, 1, { type: null, linePitchPt: null }, false, false, ctx, {})).toBe(10);
    expect(paragraphMarkLineHeight(p, 1, { type: null, linePitchPt: null }, false, true, ctx, {})).toBe(30);
  });

  it('counts an untabled 20pt East Asian mark as two cells on a 20pt grid', () => {
    const p = {
      ...para(''),
      defaultFontSize: 20,
      defaultFontFamilyEastAsia: 'ＭＳ 明朝',
    } as DocParagraph;
    let font = '';
    const ctx = {
      get font() { return font; },
      set font(v: string) { font = v; },
      measureText: () => ({
        width: 0,
        fontBoundingBoxAscent: 18,
        fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 18,
        actualBoundingBoxDescent: 2,
      } as TextMetrics),
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(
      p,
      1,
      { type: 'lines', linePitchPt: 20 },
      false,
      true,
      ctx,
      {},
    )).toBe(40);
  });
});
