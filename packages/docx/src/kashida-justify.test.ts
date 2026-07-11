import { describe, it, expect } from 'vitest';
import {
  computeKashidaDistribution,
  type KashidaLevel,
  type MeasureSegmentAdvance,
} from './kashida-justify.js';

const BEH = 'ب'; // D  Dual_Joining
const TATWEEL = 'ـ'; // U+0640 C Join_Causing

// Count U+0640 tatweels in a string.
const countTatweel = (s: string): number => [...s].filter((c) => c === TATWEEL).length;

// Mock paint model: every code point is exactly 1px wide, so an inserted tatweel
// adds exactly 1px. This makes the allocation deterministic and lets
// the tests assert measure==paint (advanceDeltaPx === #tatweels inserted).
const unitMeasure: MeasureSegmentAdvance = (_si, text) => [...text].length;

// A word of n dual-joining letters has n-1 eligible joins, of which the priority
// model chooses the final one because every join has the same Normal class.
const behWord = (n: number): string => BEH.repeat(n);

describe('computeKashidaDistribution — priority allocation', () => {
  it('fills a single Arabic word at its one priority join', () => {
    // 4 beh → 3 eligible joins; the word-final Normal tie-break chooses offset 3.
    const d = computeKashidaDistribution([{ text: behWord(4) }], 3, 'high', unitMeasure);
    expect(d).not.toBeNull();
    const plan = d!.perSeg.get(0)!;
    expect(countTatweel(plan.text)).toBe(3);
    expect(plan.advanceDeltaPx).toBe(3);
    expect(d!.appliedPx).toBe(3);
    expect(d!.residualPx).toBe(0);
    // measure==paint: the delta equals the measured growth of the augmented text.
    expect(plan.advanceDeltaPx).toBe(unitMeasure(0, plan.text) - unitMeasure(0, behWord(4)));
  });

  it('caps tatweels per selected word join (low=1, medium=2, high=unbounded)', () => {
    // 4 beh → 3 points; slack 9 is more than the points can absorb in one round.
    const run = (level: KashidaLevel) =>
      computeKashidaDistribution([{ text: behWord(4) }], 9, level, unitMeasure)!;
    const low = run('low');
    const medium = run('medium');
    const high = run('high');
    // low: 1 at the chosen join → 1 tatweel, 8px residual.
    expect(low.appliedPx).toBe(1);
    expect(low.residualPx).toBe(8);
    expect(countTatweel(low.perSeg.get(0)!.text)).toBe(1);
    // medium: 2 at the same join → 2 tatweels, 7px residual.
    expect(medium.appliedPx).toBe(2);
    expect(medium.residualPx).toBe(7);
    expect(countTatweel(medium.perSeg.get(0)!.text)).toBe(2);
    // high: fill all 9px → 9 tatweels, 0 residual.
    expect(high.appliedPx).toBe(9);
    expect(high.residualPx).toBe(0);
    expect(countTatweel(high.perSeg.get(0)!.text)).toBe(9);
  });

  it('puts all elongation for one word at its selected priority join', () => {
    // Equal Normal candidates tie toward the word end; high can repeat there.
    const d = computeKashidaDistribution([{ text: behWord(4) }], 2, 'high', unitMeasure)!;
    const plan = d!.perSeg.get(0)!;
    expect(d.appliedPx).toBe(2);
    const cps = [...plan.text];
    expect(cps.filter((c) => c === TATWEEL).length).toBe(2);
    expect(plan.text).toBe(BEH + BEH + BEH + TATWEEL.repeat(2) + BEH);
  });

  it('spreads across MULTIPLE segments before doubling up (whole-line round-robin)', () => {
    // Two 2-beh words → one point each (2 points total). slack 3, level high:
    // round1 gives each point 1 (applied 2), round2 gives the first point its 2nd
    // (applied 3). So seg0 has 2 tatweels, seg1 has 1.
    const segs = [{ text: behWord(2) }, { text: ' ' }, { text: behWord(2) }];
    const d = computeKashidaDistribution(segs, 3, 'high', unitMeasure)!;
    expect(d.appliedPx).toBe(3);
    expect(countTatweel(d.perSeg.get(0)!.text)).toBe(2);
    expect(countTatweel(d.perSeg.get(2)!.text)).toBe(1);
    // The whitespace segment is never touched.
    expect(d.perSeg.get(1)).toBeUndefined();
  });

  it('allocates scarce slack to higher-priority classes before textual order', () => {
    const d = computeKashidaDistribution([{ text: 'بب سب' }], 1, 'high', unitMeasure)!;
    const plan = d.perSeg.get(0)!;

    expect(plan.insertions).toEqual([{ beforeCp: 4, count: 1 }]);
    expect(plan.text).toBe('بب سـب');
  });

  it('returns null when the line has no eligible Arabic joining point', () => {
    expect(computeKashidaDistribution([{ text: 'abcd' }], 5, 'high', unitMeasure)).toBeNull();
    // A single beh (no interior point) and an opaque atom → null.
    expect(computeKashidaDistribution([{ text: BEH }, {}], 5, 'high', unitMeasure)).toBeNull();
  });

  it('returns null for non-positive slack (kashida never compresses)', () => {
    expect(computeKashidaDistribution([{ text: behWord(4) }], 0, 'high', unitMeasure)).toBeNull();
    expect(computeKashidaDistribution([{ text: behWord(4) }], -5, 'high', unitMeasure)).toBeNull();
  });

  it('never accepts an insertion whose measured delta is non-positive or overflows slack', () => {
    // Pathological paint model: width saturates at 5px regardless of length, so
    // after 5px of tatweels every further insertion has delta 0 and must be
    // rejected — appliedPx must not exceed the achievable growth and never exceed
    // slack.
    const capped: MeasureSegmentAdvance = (_si, text) => Math.min(5, [...text].length);
    const d = computeKashidaDistribution([{ text: behWord(4) }], 100, 'high', capped);
    // behWord(4) already measures 4 (<5); at most 1px of growth is achievable.
    expect(d).not.toBeNull();
    expect(d!.appliedPx).toBeLessThanOrEqual(1);
    expect(d!.appliedPx).toBeGreaterThan(0);
    // Paint really grew by appliedPx (measure==paint holds even under saturation).
    const plan = d!.perSeg.get(0)!;
    expect(capped(0, plan.text) - capped(0, behWord(4))).toBe(d!.appliedPx);
  });
});
