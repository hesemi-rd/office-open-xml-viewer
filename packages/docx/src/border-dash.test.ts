import { describe, expect, it } from 'vitest';
import { borderDashPattern } from './renderer.js';

// ECMA-376 §17.18.2 ST_Border dash/dot families. Patterns are in units of the
// stroked width `lw` (px); a dot = 1·lw, a dash = 3·lw, gaps = 2·lw (1·lw for
// the small-gap variant). lw=2 here so the numbers read clearly.
describe('borderDashPattern (ST_Border §17.18.2)', () => {
  const lw = 2;

  it('maps the dashed/dotted family to width-scaled patterns', () => {
    expect(borderDashPattern('dotted', lw)).toEqual([2, 4]);
    expect(borderDashPattern('dashed', lw)).toEqual([6, 4]);
    expect(borderDashPattern('dashSmallGap', lw)).toEqual([6, 2]);
    expect(borderDashPattern('dotDash', lw)).toEqual([2, 4, 6, 4]);
    expect(borderDashPattern('dotDotDash', lw)).toEqual([2, 4, 2, 4, 6, 4]);
    // dashDotStroked (thin/thick alternation) is approximated as dotDash.
    expect(borderDashPattern('dashDotStroked', lw)).toEqual([2, 4, 6, 4]);
  });

  it('scales the pattern with the border width', () => {
    expect(borderDashPattern('dashed', 1)).toEqual([3, 2]);
    expect(borderDashPattern('dashed', 4)).toEqual([12, 8]);
  });

  // Solid / continuous styles (and the separately-handled double) get no dash
  // pattern → an empty array → a continuous stroke.
  it('returns an empty pattern for solid / non-dash styles', () => {
    for (const s of ['single', 'thick', 'triple', 'double', 'wave', 'none', 'nil', 'inset']) {
      expect(borderDashPattern(s, lw)).toEqual([]);
    }
  });
});
