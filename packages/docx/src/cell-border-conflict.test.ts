import { describe, it, expect } from 'vitest';
import { resolveBorderConflict } from './cell-border-conflict.js';
import type { BorderSpec } from './types';

// ECMA-376 §17.4.66 with Word deviations from [MS-OI29500] 2.1.169 —
// adjacent table cell border conflict resolution. When cell spacing is zero,
// two cells that share an interior gridline both contribute a border for that
// edge; the DISPLAYED border is chosen by these rules (in order):
//   0. none loses to the opposing border; nil suppresses the shared edge.
//   1. a CELL border beats a TABLE border.
//   2. Word weight = width in eighth-points × Word border number; dotted and
//      dashed have weight 1 regardless of width.
//   3. equal weight ⇒ higher on the style precedence list wins.
//   4. identical style ⇒ darker colour wins, by three successive brightness
//      formulas R+B+2G, then B+2G, then G (smaller wins each).
//   5. still identical ⇒ the first in reading order (caller-provided `a`) wins.

const spec = (over: Partial<BorderSpec> = {}): BorderSpec =>
  ({ style: 'single', width: 0.5, color: null, ...over });

/** A conflicting pair: `a` is the reading-order-first candidate. Each carries a
 *  `source` flag ("cell" | "table") for rule #1. */
const cell = (over: Partial<BorderSpec> = {}) => ({ spec: spec(over), source: 'cell' as const });
const tbl = (over: Partial<BorderSpec> = {}) => ({ spec: spec(over), source: 'table' as const });

describe('§17.4.66 resolveBorderConflict', () => {
  it('Word rule 0: none loses to a real border', () => {
    expect(resolveBorderConflict(cell({ style: 'single' }), cell({ style: 'none' }))!.spec.style).toBe('single');
  });

  it('Word rule 0: nil suppresses the shared edge even against a real border', () => {
    expect(resolveBorderConflict(cell({ style: 'nil' }), cell({ style: 'single' }))).toBeNull();
    expect(resolveBorderConflict(cell({ style: 'double' }), cell({ style: 'nil' }))).toBeNull();
    expect(resolveBorderConflict(cell({ style: 'nil' }), cell({ style: 'none' }))).toBeNull();
  });

  it('one side null (absent) ⇒ the other wins', () => {
    expect(resolveBorderConflict(null, cell({ style: 'single' }))!.spec.style).toBe('single');
    expect(resolveBorderConflict(cell({ style: 'dashed' }), null)!.spec.style).toBe('dashed');
    expect(resolveBorderConflict(null, null)).toBeNull();
  });

  it('rule 1: a cell border beats a table border regardless of Word weight', () => {
    const r = resolveBorderConflict(
      tbl({ style: 'double', width: 8 }),
      cell({ style: 'single', width: 0.125 }),
    );
    expect(r!.spec.style).toBe('single');
    expect(r!.source).toBe('cell');
  });

  it('Word rule 2: width in eighth-points participates in the border weight', () => {
    // single 2pt => 16 * 1; double 0.5pt => 4 * 3.
    const r = resolveBorderConflict(
      cell({ style: 'single', width: 2 }),
      cell({ style: 'double', width: 0.5 }),
    );
    expect(r!.spec.style).toBe('single');
  });

  it('Word rule 2: dotted and dashed have weight 1 regardless of width', () => {
    const dashed = cell({ style: 'dashed', width: 20 });
    const single = cell({ style: 'single', width: 0.25 });
    expect(resolveBorderConflict(dashed, single)).toBe(single);
  });

  it('Word rule 2: uses the Word border-number table for non-dash styles', () => {
    // triple 0.125pt => 1 * 10; double 0.25pt => 2 * 3.
    const r = resolveBorderConflict(
      cell({ style: 'double', width: 0.25 }),
      cell({ style: 'triple', width: 0.125 }),
    );
    expect(r!.spec.style).toBe('triple');
  });

  it('Word rule 3: equal weight uses the documented style precedence', () => {
    // dotDash 0.125pt => 1 * 8; thick 0.5pt => 4 * 2 = 8.
    const r = resolveBorderConflict(
      cell({ style: 'dotDash', width: 0.125 }),
      cell({ style: 'thick', width: 0.5 }),
    );
    expect(r!.spec.style).toBe('thick');
  });

  it('rule 4: identical style ⇒ darker colour wins (R+B+2G smaller)', () => {
    // Same style + weight; compare brightness. Black (000000) vs red (ff0000):
    // brightness(black)=0, brightness(red)=255 ⇒ black (darker) wins.
    const r = resolveBorderConflict(cell({ style: 'single', color: '000000' }), cell({ style: 'single', color: 'ff0000' }));
    expect(r!.spec.color).toBe('000000');
  });

  it('rule 4: null (auto) colour is treated as black (0,0,0) — the darkest', () => {
    // auto ⇒ black; red loses. `a` has auto (null), `b` has red.
    const r = resolveBorderConflict(cell({ style: 'single', color: null }), cell({ style: 'single', color: 'ffffff' }));
    expect(r!.spec.color).toBeNull();
  });

  it('rule 4: brightness tie on R+B+2G broken by B+2G then G', () => {
    // Two colours with equal R+B+2G but different B+2G. Choose (R,G,B):
    //   c1 = (10, 0, 0)  ⇒ R+B+2G = 10,  B+2G = 0
    //   c2 = (0, 0, 10)  ⇒ R+B+2G = 10,  B+2G = 10
    // Equal first formula; second formula smaller for c1 ⇒ c1 wins.
    const r = resolveBorderConflict(
      cell({ style: 'single', color: '0a0000' }),
      cell({ style: 'single', color: '00000a' }),
    );
    expect(r!.spec.color).toBe('0a0000');
  });

  it('rule 5: fully identical ⇒ the reading-order-first candidate (a) wins', () => {
    const a = cell({ style: 'single', color: '808080' });
    const b = cell({ style: 'single', color: '808080' });
    const r = resolveBorderConflict(a, b);
    expect(r).toBe(a); // same reference — `a` is displayed
  });
});
