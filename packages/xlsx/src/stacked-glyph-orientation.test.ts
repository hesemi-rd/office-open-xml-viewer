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
// vo=Tr white lenticular brackets with a U+FE1x form present in the substitute
// fonts (issue #969) — still substituted upright. XLSX has no Excel ground-truth
// image; it follows the Word/PowerPoint verdict since the classifier is shared.
// The fullwidth colon ： / semicolon ； (FE13/FE14) were dropped from the substitute
// map (absent in most render fonts) and take a geometric fallback — see below.
const TR_VFORMS: Array<[string, string]> = [
  ['〖', String.fromCodePoint(0xfe17)], // left white lenticular → ︗
  ['〗', String.fromCodePoint(0xfe18)], // right white lenticular → ︘
];

interface DrawCall {
  text: string;
  x: number;
  y: number;
  /** Net canvas rotation in effect at draw time, normalised to (−π, π]. */
  rot: number;
  /** Net scale-y in effect at draw time. −1 for a reflected Tr long-stroke mark
   *  (ー 〜 ～ → `scale(1, -1)`); +1 otherwise. */
  sy: number;
  feature: string;
}

function norm(a: number): number {
  return Math.atan2(Math.sin(a), Math.cos(a));
}

function mockCtx(): { ctx: CanvasRenderingContext2D; calls: DrawCall[] } {
  let font = '20px serif';
  let textAlign: CanvasTextAlign = 'center';
  let textBaseline: CanvasTextBaseline = 'top';
  let rotation = 0;
  let sy = 1;
  const stack: Array<{ rotation: number; sy: number }> = [];
  const calls: DrawCall[] = [];
  const style = { fontFeatureSettings: 'normal' };
  const ctx = {
    canvas: { style },
    get font() { return font; }, set font(v: string) { font = v; },
    get textAlign() { return textAlign; }, set textAlign(v: CanvasTextAlign) { textAlign = v; },
    get textBaseline() { return textBaseline; }, set textBaseline(v: CanvasTextBaseline) { textBaseline = v; },
    measureText: (s: string) => ({ width: [...s].length * 20 }) as TextMetrics,
    fillText: (t: string, x: number, y: number) => calls.push({ text: t, x, y, rot: rotation, sy, feature: style.fontFeatureSettings }),
    save: () => { stack.push({ rotation, sy }); },
    restore: () => { const s = stack.pop(); if (s) { rotation = s.rotation; sy = s.sy; } },
    translate: () => {},
    rotate: (a: number) => { rotation += a; },
    scale: (_sx: number, syArg: number) => { sy *= syArg; },
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
  it('draws mirror-fallback marks upright with a per-glyph vert feature', () => {
    for (const ch of ['ー', '〜', '～']) {
      const { ctx, calls } = mockCtx();
      drawStackedVerticalChar(ctx, ch, CENTER_X, CELL_TOP, CHAR_H, true);
      expect(calls).toHaveLength(1);
      expect(calls[0].text).toBe(ch);
      expect(norm(calls[0].rot)).toBe(UPRIGHT);
      expect(calls[0].sy).toBe(1);
      expect(calls[0].feature).toBe('"vert" 1');
    }
  });

  it('keeps punctuation and brackets on their manual paths when vert is supported', () => {
    const expected: Array<[string, string, number]> = [
      ['、', '︑', UPRIGHT], ['。', '︒', UPRIGHT],
      ['：', '：', ROTATED], ['；', '；', UPRIGHT],
      ['「', '﹁', UPRIGHT], ['」', '﹂', UPRIGHT],
      ['“', '“', ROTATED], ['”', '”', ROTATED],
    ];
    for (const [ch, text, rotation] of expected) {
      const { ctx, calls } = mockCtx();
      drawStackedVerticalChar(ctx, ch, CENTER_X, CELL_TOP, CHAR_H, true);
      expect(calls).toHaveLength(1);
      expect(calls[0].text).toBe(text);
      expect(norm(calls[0].rot)).toBeCloseTo(rotation, 5);
      expect(calls[0].feature).toBe('normal');
    }
  });

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
    'SUBSTITUTES the vo=Tr white lenticular %s with its U+FE1x vertical form, upright (issue #969)',
    (orig, fe) => {
      const calls = draw(orig);
      expect(calls.length).toBe(1);
      expect(calls[0].text, `the vertical form ${fe} replaces ${orig}`).toBe(fe);
      expect(norm(calls[0].rot)).toBeCloseTo(UPRIGHT, 5);
    },
  );

  it('ROTATES the vo=Tr colon ： (geometric fallback → FE13 side-by-side dots) (issue #969)', () => {
    // FE13 is absent from most render fonts, so ：is not substituted; a 90° rotation
    // turns its two vertically-stacked dots into FE13's side-by-side dots.
    const calls = draw('：');
    expect(calls.length).toBe(1);
    expect(calls[0].text, '： is painted as its own glyph').toBe('：');
    expect(norm(calls[0].rot), '： rotates 90°').toBeCloseTo(ROTATED, 5);
  });

  it('draws the vo=Tr semicolon ； UPRIGHT (geometric fallback → FE14 dot-over-comma) (issue #969)', () => {
    // FE14 is an upright dot-over-comma (not a rotation), so ；draws upright like an
    // ideograph rather than rotating.
    const calls = draw('；');
    expect(calls.length).toBe(1);
    expect(calls[0].text, '； is painted as its own glyph').toBe('；');
    expect(norm(calls[0].rot), '； stays upright').toBeCloseTo(UPRIGHT, 5);
  });

  // The long-stroke Tr marks whose designed vertical form is the horizontal MIRROR of
  // the rotation (core verticalTrMirrorFallback): ー and the wave dash / tilde.
  it.each(['ー', '〜', '～'])(
    'ROTATES + REFLECTS the vo=Tr long-stroke mark %s by 90° plus scale(1,-1)',
    (mk) => {
      // These rotate 90° like the colon, but their font-designed vertical form is the
      // HORIZONTAL MIRROR of that rotation (Word/PowerPoint + font `vert` glyph
      // verified — a plain rotation of ー bulges LEFT, the designed form bulges RIGHT).
      // So they also reflect via `scale(1, -1)`.
      const calls = draw(mk);
      expect(calls.length).toBe(1);
      expect(calls[0].text).toBe(mk);
      expect(norm(calls[0].rot), `${mk} is rotated to a vertical bar`).toBeCloseTo(ROTATED, 5);
      expect(calls[0].sy, `${mk} is reflected (scale-y = −1)`).toBe(-1);
    },
  );

  it('does NOT reflect the vo=Tr colon ： (rotation already matches its designed form)', () => {
    // The colon's FE13 side-by-side dots fall out of the plain rotation (symmetric
    // under the mirror), so it must NOT get the scale(1,-1) reflection.
    const calls = draw('：');
    expect(calls[0].sy, '： is not reflected (scale-y = +1)').toBe(1);
  });
});
