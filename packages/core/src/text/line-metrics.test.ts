import { describe, it, expect } from 'vitest';
import { fontWinLineHeightRatio, intendedSingleLinePx, correctLineMetrics } from './line-metrics.js';

describe('fontWinLineHeightRatio', () => {
  it('returns Meiryo / Meiryo UI win line-height ratio (1.5962 em, from OS/2)', () => {
    // unitsPerEm 2048, usWinAscent 2210 + usWinDescent 1059 = 3269 → 1.5962.
    expect(fontWinLineHeightRatio('Meiryo UI')).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('Meiryo')).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('メイリオ')).toBeCloseTo(3269 / 2048, 5);
  });
  it('returns Sakkal Majalla win line-height ratio (1.3965 em, from OS/2)', () => {
    // unitsPerEm 2048, usWinAscent 1810 + usWinDescent 1050 = 2860 → 1.3965.
    expect(fontWinLineHeightRatio('Sakkal Majalla')).toBeCloseTo(2860 / 2048, 5);
    expect(fontWinLineHeightRatio('sakkal majalla')).toBeCloseTo(2860 / 2048, 5);
  });
  it('is case-insensitive', () => {
    expect(fontWinLineHeightRatio('meiryo ui')).toBeCloseTo(3269 / 2048, 5);
    expect(fontWinLineHeightRatio('MEIRYO')).toBeCloseTo(3269 / 2048, 5);
  });
  it('returns the hhea single-line ratio for tabled Latin fonts (Times New Roman, Arial)', () => {
    // Word sizes a line from the hhea line height (ascent+|descent|+lineGap), not
    // the win sum Canvas reports. Times New Roman: (1825+443+87)/2048 = 1.1499 em;
    // Arial: (1854+434+67)/2048 = 1.1499 em. Verified from the installed fonts.
    expect(fontWinLineHeightRatio('Times New Roman')).toBeCloseTo(2355 / 2048, 5);
    expect(fontWinLineHeightRatio('arial')).toBeCloseTo(2355 / 2048, 5);
  });
  it('matches Latin entries EXACTLY so variant families keep their own metrics', () => {
    // "Arial Narrow" / "Arial Black" / "Arial Nova" and any other family must NOT
    // be caught by the Arial/Times entries — they have different design metrics.
    expect(fontWinLineHeightRatio('Arial Nova')).toBeNull();
    expect(fontWinLineHeightRatio('Arial Narrow')).toBeNull();
    expect(fontWinLineHeightRatio('Arial Black')).toBeNull();
    expect(fontWinLineHeightRatio('Calibri')).toBeNull();
    expect(fontWinLineHeightRatio(null)).toBeNull();
    expect(fontWinLineHeightRatio(undefined)).toBeNull();
    expect(fontWinLineHeightRatio('')).toBeNull();
  });
});

describe('intendedSingleLinePx', () => {
  it('scales the ratio by the em size (px)', () => {
    const meiryo = 3269 / 2048;
    // 48 pt title at deviceScaleFactor 2 → em = 96 px → 1.5962 × 96.
    expect(intendedSingleLinePx('Meiryo UI', 96)).toBeCloseTo(meiryo * 96, 5);
    // Single-spaced 9 pt body at scale 2 → em = 18 px → 1.5962 × 18.
    expect(intendedSingleLinePx('Meiryo UI', 18)).toBeCloseTo(meiryo * 18, 5);
  });
  it('returns 0 (no-op sentinel) for untabled fonts', () => {
    expect(intendedSingleLinePx('Calibri', 96)).toBe(0);
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
  it('keeps the substitute metrics when its box is SMALLER than the document win box (Meiryo)', () => {
    // Two-regime rule: a substitute that UNDERSTATES the document font (here
    // 18px natural vs Meiryo's 1.5962em ≈ 28.7px) passes through unchanged —
    // the intendedSingleLinePx floor raises the LINE BOX and the renderer
    // centers the natural line, keeping ink where Word's sits (sample-3 VRT).
    const r = correctLineMetrics('Meiryo UI', 18, 14, 4);
    expect(r).toEqual({ ascent: 14, descent: 4 });
    // ...while the floor still claims the document font's win height.
    expect(intendedSingleLinePx('Meiryo UI', 18)).toBeCloseTo((3269 / 2048) * 18, 5);
  });
  it('passes through measured metrics unchanged for untabled fonts', () => {
    const r = correctLineMetrics('Calibri', 12, 11, 3);
    expect(r).toEqual({ ascent: 11, descent: 3 });
  });
  it('keeps the measured box for an installed Latin font shorter than its design box', () => {
    // Times New Roman is tabled (design 1.1499 em). At em = 12 the design box is
    // ~13.8 px; the Canvas win box (≈1.107 em ≈ 13.3 px, here 10.7 + 2.6) is
    // SMALLER, so correctLineMetrics passes it through and the intendedSingleLinePx
    // floor (not this function) raises the LINE BOX — matching the Meiryo regime.
    const r = correctLineMetrics('Times New Roman', 12, 10.7, 2.6);
    expect(r).toEqual({ ascent: 10.7, descent: 2.6 });
    expect(intendedSingleLinePx('Times New Roman', 12)).toBeCloseTo((2355 / 2048) * 12, 5);
  });
});
