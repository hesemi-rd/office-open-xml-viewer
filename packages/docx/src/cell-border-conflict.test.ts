import { describe, it, expect } from 'vitest';
import { resolveBorderConflict } from './cell-border-conflict.js';
import type { BorderSpec } from './types';

// ECMA-376 §17.4.66 — adjacent table cell border conflict resolution. When cell
// spacing is zero, two cells that share an interior gridline both contribute a
// border for that edge; the DISPLAYED border is chosen by these rules (in order):
//   0. nil/none loses to any real border (if BOTH nil ⇒ nothing).
//   1. a CELL border beats a TABLE border.
//   2. weight = (# of lines in the style) × (style rank number); larger wins.
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
  it('rule 0: nil/none loses to a real border', () => {
    expect(resolveBorderConflict(cell({ style: 'nil' }), cell({ style: 'single' }))!.spec.style).toBe('single');
    expect(resolveBorderConflict(cell({ style: 'single' }), cell({ style: 'none' }))!.spec.style).toBe('single');
  });

  it('rule 0: both nil/none ⇒ no border (null)', () => {
    expect(resolveBorderConflict(cell({ style: 'nil' }), cell({ style: 'none' }))).toBeNull();
  });

  it('one side null (absent) ⇒ the other wins', () => {
    expect(resolveBorderConflict(null, cell({ style: 'single' }))!.spec.style).toBe('single');
    expect(resolveBorderConflict(cell({ style: 'dashed' }), null)!.spec.style).toBe('dashed');
    expect(resolveBorderConflict(null, null)).toBeNull();
  });

  it('rule 1: a cell border beats a table border regardless of weight', () => {
    // table `double` (weight 2×3=6) vs cell `single` (1×1=1): the cell still wins.
    const r = resolveBorderConflict(tbl({ style: 'double' }), cell({ style: 'single' }));
    expect(r!.spec.style).toBe('single');
    expect(r!.source).toBe('cell');
  });

  it('rule 2: heavier weight wins (double 2×3=6 beats single 1×1=1)', () => {
    const r = resolveBorderConflict(cell({ style: 'single' }), cell({ style: 'double' }));
    expect(r!.spec.style).toBe('double');
  });

  it('rule 2: dashed (1×5=5) beats thick (1×2=2)', () => {
    const r = resolveBorderConflict(cell({ style: 'thick' }), cell({ style: 'dashed' }));
    expect(r!.spec.style).toBe('dashed');
  });

  it('rule 2: triple (3×8=24) beats double (2×3=6)', () => {
    const r = resolveBorderConflict(cell({ style: 'double' }), cell({ style: 'triple' }));
    expect(r!.spec.style).toBe('triple');
  });

  it('rule 3: equal weight ⇒ higher on the precedence list wins (single before thick)', () => {
    // Contrive equal weight: single (1×1=1) vs … there is no other weight-1 style,
    // so use two styles that tie. thinThickSmallGap (2×9=18) vs … pick a real tie:
    // dotDash (1×6=6) vs double (2×3=6) — equal weight 6. Precedence: double is
    // higher on the list than dotDash, so double wins.
    const r = resolveBorderConflict(cell({ style: 'dotDash' }), cell({ style: 'double' }));
    expect(r!.spec.style).toBe('double');
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
