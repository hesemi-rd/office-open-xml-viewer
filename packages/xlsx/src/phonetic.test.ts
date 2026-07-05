import { describe, it, expect } from 'vitest';
import { placePhoneticRuns, toCodePoints } from './phonetic.js';
import type { PhoneticRun } from './types.js';

/** A fixed-advance measurer: every code point is `w` wide. Lets us assert exact
 *  band geometry without a Canvas. */
function fixedMeasure(w: number) {
  return (s: string) => toCodePoints(s).length * w;
}

describe('placePhoneticRuns (§18.4.6 / §18.18.56)', () => {
  const base = '課長'; // 2 code points
  const runs: PhoneticRun[] = [
    { sb: 0, eb: 1, text: 'カ' },
    { sb: 1, eb: 2, text: 'チョウ' },
  ];

  it('left alignment: each hint band starts at its base span left edge', () => {
    // base char width 10, baseLeftX 100 → char0 span [100,110), char1 [110,120)
    const placed = placePhoneticRuns(runs, base, 100, 'left', fixedMeasure(10));
    expect(placed).toEqual([
      { text: 'カ', x: 100, width: 10, spread: 'start' },
      { text: 'チョウ', x: 110, width: 10, spread: 'start' },
    ]);
  });

  it('center alignment: band spans the base char, spread=center', () => {
    const placed = placePhoneticRuns(runs, base, 0, 'center', fixedMeasure(10));
    expect(placed).toEqual([
      { text: 'カ', x: 0, width: 10, spread: 'center' },
      { text: 'チョウ', x: 10, width: 10, spread: 'center' },
    ]);
  });

  it('distributed alignment: spread=distribute over the base span width', () => {
    const placed = placePhoneticRuns(runs, base, 0, 'distributed', fixedMeasure(10));
    expect(placed.map((p) => p.spread)).toEqual(['distribute', 'distribute']);
  });

  it('multi-char base span sums advances for the start offset', () => {
    // base "受注実績" (4 cps), a run over chars [2,4)
    const b = '受注実績';
    const r: PhoneticRun[] = [{ sb: 2, eb: 4, text: 'ジッセキ' }];
    const placed = placePhoneticRuns(r, b, 0, 'left', fixedMeasure(10));
    // span start = 2 chars × 10 = 20; width = 2 chars × 10 = 20
    expect(placed).toEqual([{ text: 'ジッセキ', x: 20, width: 20, spread: 'start' }]);
  });

  it('skips runs with sb >= eb or out of range (defensive, spec requires sb<eb)', () => {
    const bad: PhoneticRun[] = [
      { sb: 1, eb: 1, text: 'x' }, // empty span
      { sb: 5, eb: 6, text: 'y' }, // out of range (base only 2 cps)
      { sb: 0, eb: 1, text: 'ok' },
    ];
    const placed = placePhoneticRuns(bad, base, 0, 'left', fixedMeasure(10));
    expect(placed).toEqual([{ text: 'ok', x: 0, width: 10, spread: 'start' }]);
  });

  it('clamps eb to the base length', () => {
    const r: PhoneticRun[] = [{ sb: 0, eb: 99, text: 'z' }];
    const placed = placePhoneticRuns(r, base, 0, 'left', fixedMeasure(10));
    // eb clamps to 2 → width = 2 × 10 = 20
    expect(placed).toEqual([{ text: 'z', x: 0, width: 20, spread: 'start' }]);
  });

  it('counts a surrogate pair as one base character', () => {
    // "𠀋" (U+2000B) is a single code point but two UTF-16 units.
    const b = '𠀋田'; // 2 code points
    const r: PhoneticRun[] = [{ sb: 1, eb: 2, text: 'タ' }];
    const placed = placePhoneticRuns(r, b, 0, 'left', fixedMeasure(10));
    // char1 (田) span start = 1 cp × 10 = 10, not 2 (would be wrong with .length)
    expect(placed).toEqual([{ text: 'タ', x: 10, width: 10, spread: 'start' }]);
  });
});
