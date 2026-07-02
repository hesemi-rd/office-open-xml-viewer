import { describe, it, expect } from 'vitest';
import { elideToWidth } from './text-elide.js';

/**
 * Mock canvas context whose `measureText` returns a width proportional to the
 * summed per-character weight. A weight map lets a test model variable-width
 * fonts (e.g. CJK glyphs twice as wide as Latin) so we can prove the elision is
 * width-based, not character-count-based.
 */
function mockCtx(perChar: number, weights?: Record<string, number>): CanvasRenderingContext2D {
  return {
    measureText(s: string) {
      let w = 0;
      for (const ch of s) w += (weights?.[ch] ?? 1) * perChar;
      return { width: w } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

describe('elideToWidth', () => {
  it('returns the text unchanged when it already fits', () => {
    const ctx = mockCtx(10); // each char = 10px
    // "abc" = 30px, fits in 30
    expect(elideToWidth(ctx, 'abc', 30)).toBe('abc');
    expect(elideToWidth(ctx, 'abc', 100)).toBe('abc');
  });

  it('appends an ellipsis and keeps the longest fitting prefix', () => {
    const ctx = mockCtx(10); // char = 10px, ellipsis '…' = 10px
    // "abcdef" = 60px, does not fit in 45.
    // prefix "abc" + "…" = 40 ≤ 45; "abcd" + "…" = 50 > 45 → keep "abc".
    expect(elideToWidth(ctx, 'abcdef', 45)).toBe('abc…');
  });

  it('boundary: max exactly fits prefix+ellipsis', () => {
    const ctx = mockCtx(10);
    // "ab" + "…" = 30 == 30 → fits; "abc" + "…" = 40 > 30.
    expect(elideToWidth(ctx, 'abcdef', 30)).toBe('ab…');
  });

  it('boundary: full string width exactly equals maxPx → no ellipsis', () => {
    const ctx = mockCtx(10);
    expect(elideToWidth(ctx, 'abcd', 40)).toBe('abcd');
    // one pixel short → must elide. char=ellipsis=10px, so a fitting prefix p
    // needs 10·len(p)+10 ≤ 39 → len(p) ≤ 2 → "ab…".
    expect(elideToWidth(ctx, 'abcd', 39)).toBe('ab…');
  });

  it('returns bare ellipsis when only the ellipsis fits', () => {
    const ctx = mockCtx(10);
    // ellipsis = 10px; one char + ellipsis = 20 > 15 → just "…".
    expect(elideToWidth(ctx, 'abcdef', 15)).toBe('…');
    expect(elideToWidth(ctx, 'abcdef', 10)).toBe('…');
  });

  it('returns empty string when not even the ellipsis fits', () => {
    const ctx = mockCtx(10); // ellipsis = 10px
    expect(elideToWidth(ctx, 'abc', 9)).toBe('');
    expect(elideToWidth(ctx, 'abc', 0)).toBe('');
    expect(elideToWidth(ctx, 'abc', -5)).toBe('');
  });

  it('empty text returns empty', () => {
    const ctx = mockCtx(10);
    expect(elideToWidth(ctx, '', 100)).toBe('');
    expect(elideToWidth(ctx, '', 0)).toBe('');
  });

  it('is width-based, not char-count-based (CJK glyphs are wider)', () => {
    // Latin chars = 8px, CJK chars = 16px, ellipsis '…' = 8px.
    const ctx = mockCtx(8, { '年': 2, '月': 2, '期': 2, '…': 1 });
    // "2025年3月期" widths: '2'8 '0'8 '2'8 '5'8 '年'16 '3'8 '月'16 '期'16 = 88px.
    // Fits in 90 unchanged.
    expect(elideToWidth(ctx, '2025年3月期', 90)).toBe('2025年3月期');
    // In 60px: find longest prefix p with width(p)+8 ≤ 60 → width(p) ≤ 52.
    //   "2025年" = 8+8+8+8+16 = 48; +8 = 56 ≤ 60 ✓
    //   "2025年3" = 56; +8 = 64 > 60 ✗  → keep "2025年".
    expect(elideToWidth(ctx, '2025年3月期', 60)).toBe('2025年…');
    // A pure-Latin string of the SAME char count fits far more before eliding,
    // proving the cut point depends on measured width, not length.
    // "20253xyz" (8 chars, all 8px) = 64px; in 60 → prefix ≤ 52 → 6 chars "20253x"+… = 56.
    expect(elideToWidth(ctx, '20253xyz', 60)).toBe('20253x…');
  });

  it('narrow slot keeps only what fits for a long CJK label', () => {
    const ctx = mockCtx(10, { '…': 1 }); // CJK-ish: every char 10px, ellipsis 10px
    // 5 wide chars = 50px; in 25px → prefix p with 10·len(p)+10 ≤ 25 → len ≤ 1.
    expect(elideToWidth(ctx, '一二三四五', 25)).toBe('一…');
  });
});
