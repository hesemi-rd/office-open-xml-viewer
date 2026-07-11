import { describe, it, expect } from 'vitest';
import { drawStackedVerticalChar } from './vertical-text.js';

// ECMA-376 §18.8.1 alignment `textRotation="255"` — Excel "stacked" vertical
// text: each character is stacked upright, one per line, top to bottom. Before
// this fix every character was drawn upright with no orientation logic, so:
//   • vo=Tr fullwidth brackets （）「」… kept their HORIZONTAL shape (Excel shows
//     the vertical presentation form), and
//   • vo=Tr ー (U+30FC prolonged sound mark) stayed a horizontal dash instead of
//     the vertical bar Excel draws, and
//   • vo=Tu commas / full stops 、。 were not swapped for their upper-right
//     corner-hanging vertical form.
// The fix routes each stacked glyph through the core UAX#50 classifier:
//   • vo=U  (CJK / kana / fullwidth digits) → upright, unchanged (Excel stacks
//     these upright).
//   • vo=R  (Latin / ASCII digits) → upright, unchanged (Excel stacks Latin
//     upright too — NOT rotated, unlike pptx/docx eaVert).
//   • vo=Tu with a U+FE1x form (、。，) → SUBSTITUTE that vertical form (upright).
//   • vo=Tr with a U+FE3x form (（）「」…) → SUBSTITUTE that vertical form (upright,
//     "substitute-first" per UAX#50 §5).
//   • vo=Tr with NO vertical form (ー, quotes) → ROTATE the glyph 90° (the Tr
//     fallback).

const CENTER_X = 50;
const CELL_TOP = 10;
const CHAR_H = 20;

const U_CJK = '国';
const R_LATIN = 'A';
const TR_BRACKET = '（';
const TR_BRACKET_FE = String.fromCodePoint(0xfe35); // ︵
const TU_COMMA = '、';
const TU_COMMA_FE = String.fromCodePoint(0xfe11);
const TR_ROTATE = 'ー'; // U+30FC — no vertical form → rotate
// vo=Tr punctuation / white lenticular brackets with a U+FE1x form (issue #969).
// XLSX has no Excel ground-truth image; it follows the Word/PowerPoint verdict
// (docx tbRl + pptx eaVert, PDF-adjudicated) since the classifier is shared.
const TR_VFORMS: Array<[string, string]> = [
  ['：', String.fromCodePoint(0xfe13)], // fullwidth colon → ︓
  ['；', String.fromCodePoint(0xfe14)], // fullwidth semicolon → ︔
  ['〖', String.fromCodePoint(0xfe17)], // left white lenticular → ︗
  ['〗', String.fromCodePoint(0xfe18)], // right white lenticular → ︘
];

interface DrawCall {
  text: string;
  x: number;
  y: number;
  /** Net canvas rotation in effect at draw time, normalised to (−π, π]. */
  rot: number;
}

function norm(a: number): number {
  return Math.atan2(Math.sin(a), Math.cos(a));
}

function mockCtx(): { ctx: CanvasRenderingContext2D; calls: DrawCall[] } {
  let font = '20px serif';
  let textAlign: CanvasTextAlign = 'center';
  let textBaseline: CanvasTextBaseline = 'top';
  let rotation = 0;
  const stack: number[] = [];
  const calls: DrawCall[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get textAlign() { return textAlign; }, set textAlign(v: CanvasTextAlign) { textAlign = v; },
    get textBaseline() { return textBaseline; }, set textBaseline(v: CanvasTextBaseline) { textBaseline = v; },
    measureText: (s: string) => ({ width: [...s].length * 20 }) as TextMetrics,
    fillText: (t: string, x: number, y: number) => calls.push({ text: t, x, y, rot: rotation }),
    save: () => { stack.push(rotation); },
    restore: () => { rotation = stack.pop() ?? rotation; },
    translate: () => {},
    rotate: (a: number) => { rotation += a; },
    fillStyle: '#000',
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function draw(ch: string): DrawCall[] {
  const { ctx, calls } = mockCtx();
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  drawStackedVerticalChar(ctx, ch, CENTER_X, CELL_TOP, CHAR_H);
  return calls;
}

const UPRIGHT = 0;
const ROTATED = Math.PI / 2;

describe('xlsx stacked text — UAX#50 per-glyph orientation (textRotation=255, issue #790)', () => {
  it('keeps vo=U CJK upright (unchanged)', () => {
    const calls = draw(U_CJK);
    expect(calls.length).toBe(1);
    expect(calls[0].text).toBe(U_CJK);
    expect(norm(calls[0].rot)).toBeCloseTo(UPRIGHT, 5);
  });

  it('keeps vo=R Latin upright (Excel stacks Latin upright, not rotated)', () => {
    const calls = draw(R_LATIN);
    expect(calls.length).toBe(1);
    expect(calls[0].text).toBe(R_LATIN);
    expect(norm(calls[0].rot)).toBeCloseTo(UPRIGHT, 5);
  });

  it('SUBSTITUTES a vo=Tr fullwidth bracket with its vertical form, upright', () => {
    const calls = draw(TR_BRACKET);
    expect(calls.length).toBe(1);
    expect(calls[0].text, 'the U+FE35 vertical form ︵ replaces （').toBe(TR_BRACKET_FE);
    expect(norm(calls[0].rot)).toBeCloseTo(UPRIGHT, 5);
  });

  it('SUBSTITUTES a vo=Tu comma with its U+FE11 vertical form, upright', () => {
    const calls = draw(TU_COMMA);
    expect(calls.length).toBe(1);
    expect(calls[0].text).toBe(TU_COMMA_FE);
    expect(norm(calls[0].rot)).toBeCloseTo(UPRIGHT, 5);
  });

  it.each(TR_VFORMS)(
    'SUBSTITUTES the vo=Tr colon/semicolon/lenticular %s with its U+FE1x vertical form, upright (issue #969)',
    (orig, fe) => {
      const calls = draw(orig);
      expect(calls.length).toBe(1);
      expect(calls[0].text, `the vertical form ${fe} replaces ${orig}`).toBe(fe);
      expect(norm(calls[0].rot)).toBeCloseTo(UPRIGHT, 5);
    },
  );

  it('ROTATES a vo=Tr glyph with no vertical form (ー) by 90°', () => {
    const calls = draw(TR_ROTATE);
    expect(calls.length).toBe(1);
    expect(calls[0].text).toBe(TR_ROTATE);
    expect(norm(calls[0].rot), 'ー is rotated to a vertical bar').toBeCloseTo(ROTATED, 5);
  });
});
