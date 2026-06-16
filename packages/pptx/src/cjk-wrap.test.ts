import { describe, it, expect } from 'vitest';
import { DEFAULT_KINSOKU_RULES, resolveKinsokuRules } from '@silurus/ooxml-core';
import { fitCjkLine, type MeasuredChar } from './cjk-wrap.js';

const mc = (s: string, w = 10): MeasuredChar[] => [...s].map((ch) => ({ ch, w }));

describe('fitCjkLine — 行頭禁則 (line-start-forbidden)', () => {
  it('retracts so a line never begins with 、', () => {
    // width fits "あい" (20px); naive split would start the next line with 、.
    // kinsoku pulls い down: head = "あ" (1 char).
    expect(fitCjkLine(mc('あい、う'), 0, 20, DEFAULT_KINSOKU_RULES)).toBe(1);
  });

  it('retracts so a line never begins with 。', () => {
    expect(fitCjkLine(mc('かき。く'), 0, 20, DEFAULT_KINSOKU_RULES)).toBe(1);
  });
});

describe('fitCjkLine — 行末禁則 (line-end-forbidden)', () => {
  it('does not leave an opening bracket 「 dangling at line end', () => {
    // width fits "あ「" (20px); 「 may not end a line → pull it down.
    expect(fitCjkLine(mc('あ「い'), 0, 20, DEFAULT_KINSOKU_RULES)).toBe(1);
  });
});

describe('fitCjkLine — legal splits & progress', () => {
  it('leaves a kinsoku-legal greedy split untouched', () => {
    expect(fitCjkLine(mc('あいうえ'), 0, 20, DEFAULT_KINSOKU_RULES)).toBe(2);
  });

  it('returns the whole run when it all fits', () => {
    expect(fitCjkLine(mc('あい'), 0, 100, DEFAULT_KINSOKU_RULES)).toBe(2);
  });

  it('places at least one char on an empty line even if it overflows', () => {
    expect(fitCjkLine([{ ch: 'あ', w: 30 }, { ch: 'い', w: 10 }], 0, 20, DEFAULT_KINSOKU_RULES)).toBe(1);
  });

  it('returns 0 when nothing fits on a non-empty line (caller breaks first)', () => {
    // line already holds 15px; first char (10px) would overflow → push run down.
    expect(fitCjkLine(mc('、あ'), 15, 20, DEFAULT_KINSOKU_RULES)).toBe(0);
  });

  it('disabled kinsoku returns the plain greedy split', () => {
    const off = resolveKinsokuRules({ kinsoku: false });
    expect(fitCjkLine(mc('あい、う'), 0, 20, off)).toBe(2);
  });

  it('treats an astral (surrogate-pair) grapheme as one unit', () => {
    // [...'𠮷𠮷、'] yields 3 code-point graphemes; width fits 2 (𠮷𠮷), and
    // kinsoku pulls the 2nd 𠮷 down so the line never begins with 、.
    expect(fitCjkLine(mc('𠮷𠮷、'), 0, 20, DEFAULT_KINSOKU_RULES)).toBe(1);
  });

  it('falls back to the greedy split for a pathological all-forbidden run', () => {
    // Every char is 行頭禁則; no legal split exists within minSplit, so the
    // greedy split (2) is returned rather than hanging or emptying the line.
    expect(fitCjkLine(mc('、。、。'), 0, 20, DEFAULT_KINSOKU_RULES)).toBe(2);
  });
});
