import { describe, it, expect } from 'vitest';
import { fontWinLineHeightRatio, intendedSingleLinePx, correctLineMetrics } from './font-metrics.js';

describe('fontWinLineHeightRatio', () => {
  it('returns Meiryo / Meiryo UI win line-height ratio (≈1.60 em)', () => {
    expect(fontWinLineHeightRatio('Meiryo UI')).toBe(1.6);
    expect(fontWinLineHeightRatio('Meiryo')).toBe(1.6);
    expect(fontWinLineHeightRatio('メイリオ')).toBe(1.6);
  });
  it('returns Sakkal Majalla win line-height ratio (1.3965 em, from OS/2)', () => {
    // unitsPerEm 2048, usWinAscent 1810 + usWinDescent 1050 = 2860 → 1.3965.
    expect(fontWinLineHeightRatio('Sakkal Majalla')).toBeCloseTo(2860 / 2048, 5);
    expect(fontWinLineHeightRatio('sakkal majalla')).toBeCloseTo(2860 / 2048, 5);
  });
  it('is case-insensitive', () => {
    expect(fontWinLineHeightRatio('meiryo ui')).toBe(1.6);
    expect(fontWinLineHeightRatio('MEIRYO')).toBe(1.6);
  });
  it('returns null for untabled fonts (Latin / unknown / null)', () => {
    // Latin fonts are intentionally absent — their win ratio (~1.15–1.22) is
    // close to the browser fallback, so no correction is needed.
    expect(fontWinLineHeightRatio('Arial')).toBeNull();
    expect(fontWinLineHeightRatio('Calibri')).toBeNull();
    expect(fontWinLineHeightRatio('Arial Nova')).toBeNull();
    expect(fontWinLineHeightRatio(null)).toBeNull();
    expect(fontWinLineHeightRatio(undefined)).toBeNull();
    expect(fontWinLineHeightRatio('')).toBeNull();
  });
});

describe('intendedSingleLinePx', () => {
  it('scales the ratio by the em size (px)', () => {
    // 48 pt title at deviceScaleFactor 2 → em = 96 px → 1.6 × 96 = 153.6.
    expect(intendedSingleLinePx('Meiryo UI', 96)).toBeCloseTo(153.6, 5);
    // Single-spaced 9 pt body at scale 2 → em = 18 px → 1.6 × 18 = 28.8.
    expect(intendedSingleLinePx('Meiryo UI', 18)).toBeCloseTo(28.8, 5);
  });
  it('returns 0 (no-op sentinel) for untabled fonts', () => {
    expect(intendedSingleLinePx('Arial', 96)).toBe(0);
    expect(intendedSingleLinePx(null, 96)).toBe(0);
  });
});

describe('correctLineMetrics', () => {
  it('returns the document font win ascent/descent for tabled fonts', () => {
    // Sakkal Majalla at em = 12 px: asc = 1810/2048 × 12, desc = 1050/2048 × 12.
    // The substitute's (over-large) measured metrics are replaced, not scaled.
    const r = correctLineMetrics('Sakkal Majalla', 12, /*substituteAsc*/ 18, /*substituteDesc*/ 8);
    expect(r.ascent).toBeCloseTo((1810 / 2048) * 12, 5);
    expect(r.descent).toBeCloseTo((1050 / 2048) * 12, 5);
    // Total equals the win line-height ratio × em (here ~16.76 px, well under
    // the substitute's 26 px), which is what fixes the over-measured cell box.
    expect(r.ascent + r.descent).toBeCloseTo((2860 / 2048) * 12, 5);
  });
  it('preserves Meiryo total at exactly 1.60 em with the OS/2 split', () => {
    const r = correctLineMetrics('Meiryo UI', 18, 14, 4);
    expect(r.ascent + r.descent).toBeCloseTo(1.6 * 18, 5);
    expect(r.ascent / (r.ascent + r.descent)).toBeCloseTo(2210 / 3269, 5);
  });
  it('passes through measured metrics unchanged for untabled fonts', () => {
    const r = correctLineMetrics('Arial', 12, 11, 3);
    expect(r).toEqual({ ascent: 11, descent: 3 });
  });
});
