import { describe, it, expect } from 'vitest';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// A POSITIVE first-line indent (a:pPr@indent > 0 on a NON-bullet paragraph)
// shifts the first line right at draw time AND narrows that line's alignment /
// justify region by the indent. The wrap/layout pass must reduce the FIRST
// line's wrap budget by the same amount, otherwise the first line wraps too
// late and overruns the right margin (marR). Continuation lines keep the full
// width. This mirrors willTextOverflow's `textMaxW - firstLineIndent` and the
// draw-side `textMaxW - textXOffset`, and matches xlsx PR #620.

function mockCtx() {
  let font = '';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    // Deterministic: every glyph (incl. spaces) is 10px wide.
    measureText: (s: string) => ({ width: s.length * 10 }),
    fillRect() {}, fillText() {},
    fillStyle: '', strokeStyle: '',
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial', ...over,
  };
}

function para(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null,
    defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
}

function lineTexts(line: { segments: { text: string }[] }): string {
  return line.segments.map((s) => s.text).join('');
}

// Width a line occupies = total glyph count × 10 (mock), trailing whitespace
// included — matching how `lineW` accumulates inside layoutParagraph.
function lineWidth(line: { segments: { text: string }[] }): number {
  return lineTexts(line).length * 10;
}

// Count the non-whitespace tokens that begin the first line.
function firstLineWordCount(line: { segments: { text: string }[] }): number {
  return lineTexts(line).trim().split(/\s+/).filter((t) => t.length > 0).length;
}

describe('pptx non-bullet first-line indent narrows the first wrap line', () => {
  // Five 4-char words, single spaces between: "aaaa bbbb cccc dddd eeee".
  // Each word = 40px, each space = 10px. maxWidth = 200px.
  const TEXT = 'aaaa bbbb cccc dddd eeee';
  const MAX = 200;

  it('packs more tokens on the first line when there is no indent (baseline)', () => {
    const lines = layoutParagraph(
      mockCtx(), para([run(TEXT)]), MAX, 20, '000000', 1, 0,
    );
    // "aaaa bbbb cccc dddd" = 4*40 + 3*10 = 190 ≤ 200; "+ eeee" would be 240 > 200.
    expect(firstLineWordCount(lines[0])).toBe(4);
    expect(lineWidth(lines[0])).toBeLessThanOrEqual(MAX);
  });

  it('packs FEWER tokens on the first line with a positive first-line indent', () => {
    // indent of 80px (~2 words). First-line budget becomes 200 - 80 = 120.
    // "aaaa bbbb" = 90 ≤ 120; "+ cccc" = 140 > 120, so only 2 words fit.
    const indentPx = 80;
    const lines = layoutParagraph(
      mockCtx(), para([run(TEXT)]), MAX, 20, '000000', 1, 0,
      false, false, 1.0, undefined,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      indentPx,
    );
    // First line wraps earlier than the baseline (4 → 2 words).
    expect(firstLineWordCount(lines[0])).toBeLessThan(4);
    // No overrun: first line width ≤ maxWidth - indent.
    expect(lineWidth(lines[0])).toBeLessThanOrEqual(MAX - indentPx);
    // Continuation lines still use the FULL width (not the narrowed budget):
    // the second line should be allowed to exceed (MAX - indentPx).
    const continuationMax = Math.max(...lines.slice(1).map(lineWidth));
    expect(continuationMax).toBeGreaterThan(MAX - indentPx);
    expect(continuationMax).toBeLessThanOrEqual(MAX);
  });

  it('is byte-identical to the default when firstLineIndentPx is 0 or omitted', () => {
    const omitted = layoutParagraph(
      mockCtx(), para([run(TEXT)]), MAX, 20, '000000', 1, 0,
    );
    const explicitZero = layoutParagraph(
      mockCtx(), para([run(TEXT)]), MAX, 20, '000000', 1, 0,
      false, false, 1.0, undefined,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      0,
    );
    const texts = (ls: { segments: { text: string }[] }[]) => ls.map(lineTexts);
    expect(texts(explicitZero)).toEqual(texts(omitted));
  });
});
