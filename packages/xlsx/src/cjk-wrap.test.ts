import { describe, it, expect } from 'vitest';
import { wrapParagraphLines, layoutRichTextLines } from './renderer.js';
import type { CellFont, Run } from './types.js';

/** Stub 2D context: every grapheme measures 10px wide, so maxWidth=30 fits
 *  exactly 3 CJK chars. No real canvas / font loading required. */
const ctx = {
  font: '',
  measureText: (s: string) => ({ width: [...s].length * 10 }),
} as unknown as CanvasRenderingContext2D;

const baseFont: CellFont = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: null,
};

const richText = (lines: ReturnType<typeof layoutRichTextLines>): string[] =>
  lines.map((l) => l.segments.map((s) => s.text).join(''));

describe('wrapParagraphLines — 行頭禁則 (line-start-forbidden)', () => {
  it('never starts a wrapped line with 、 (pulls the preceding char down)', () => {
    // Naive greedy at maxWidth=30: ["あいう", "、え"] → line 2 starts with 、.
    // Kinsoku retracts: "う" moves down so the break is legal.
    const out = wrapParagraphLines(ctx, 'あいう、え', 30);
    expect(out.every((line) => !line.startsWith('、'))).toBe(true);
    expect(out).toEqual(['あい', 'う、え']);
  });

  it('never starts a wrapped line with 。', () => {
    const out = wrapParagraphLines(ctx, 'かきく。け', 30);
    expect(out.every((line) => !line.startsWith('。'))).toBe(true);
    expect(out).toEqual(['かき', 'く。け']);
  });
});

describe('wrapParagraphLines — 行末禁則 (line-end-forbidden)', () => {
  it('never ends a wrapped line with 「 (pushes the opener down)', () => {
    // Naive greedy at maxWidth=30: ["あい「", "うえ"] → line 1 ends with 「.
    // Kinsoku retracts: "「" moves down to lead line 2.
    const out = wrapParagraphLines(ctx, 'あい「うえ', 30);
    expect(out.every((line) => !line.endsWith('「'))).toBe(true);
    expect(out).toEqual(['あい', '「うえ']);
  });
});

describe('wrapParagraphLines — parity (no forbidden chars)', () => {
  it('wraps plain CJK identically to the pre-kinsoku greedy result', () => {
    // No forbidden chars at any boundary → behaviour must be unchanged.
    expect(wrapParagraphLines(ctx, 'あいうえお', 30)).toEqual(['あいう', 'えお']);
  });

  it('wraps a longer plain CJK run identically', () => {
    expect(wrapParagraphLines(ctx, 'あいうえおかきくけこ', 30)).toEqual([
      'あいう',
      'えおか',
      'きくけ',
      'こ',
    ]);
  });

  it('does not retract when the break is already legal', () => {
    // The 、 lands at the END of line 1, which is allowed; no change.
    expect(wrapParagraphLines(ctx, 'あい、うえ', 30)).toEqual(['あい、', 'うえ']);
  });
});

describe('layoutRichTextLines — 行頭禁則 (line-start-forbidden)', () => {
  const cjkRun = (text: string): Run[] => [{ text }];

  it('never starts a wrapped line with 、 across the segment break', () => {
    const lines = layoutRichTextLines(ctx, cjkRun('あいう、え'), baseFont, 1, 30);
    const texts = richText(lines);
    expect(texts.every((t) => !t.startsWith('、'))).toBe(true);
    expect(texts).toEqual(['あい', 'う、え']);
  });

  it('never ends a wrapped line with 「', () => {
    const lines = layoutRichTextLines(ctx, cjkRun('あい「うえ'), baseFont, 1, 30);
    const texts = richText(lines);
    expect(texts.every((t) => !t.endsWith('「'))).toBe(true);
    expect(texts).toEqual(['あい', '「うえ']);
  });

  it('retracts across a run/segment boundary (forbidden char in its own run)', () => {
    // The 行頭禁則 char 、 is the first code point of the SECOND run, and the
    // char to pull down ("う") is the last segment of the FIRST run. The
    // retracted segment is whole, so it is moved (not split) onto the next line.
    const runs: Run[] = [{ text: 'あいう' }, { text: '、え' }];
    const lines = layoutRichTextLines(ctx, runs, baseFont, 1, 30);
    const texts = richText(lines);
    expect(texts.every((t) => !t.startsWith('、'))).toBe(true);
    expect(texts).toEqual(['あい', 'う、え']);
  });
});

describe('layoutRichTextLines — parity (no forbidden chars)', () => {
  it('wraps plain CJK identically to the pre-kinsoku greedy result', () => {
    const lines = layoutRichTextLines(ctx, [{ text: 'あいうえお' }], baseFont, 1, 30);
    expect(richText(lines)).toEqual(['あいう', 'えお']);
  });
});
