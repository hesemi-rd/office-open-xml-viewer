import { describe, expect, it } from 'vitest';
import { doubleRailGeometry, fillDoubleBorder } from './double-border.js';

// ECMA-376 §17.18.2 / §18.18.3 "double": two parallel lines with a gap, laid out
// in device pixels (lw·dpr) as three floored-thirds bands rail / gap / rail.
describe('doubleRailGeometry', () => {
  // THE regression gate: a thin double (e.g. a table "Total" row top at
  // sz6 ≈ 0.75px) must never render the two rails overlapping. For ANY width/dpr
  // the rails stay ≥ 1 device px apart (gapDev ≥ 1) and each rail paints
  // (railDev ≥ 1), so the double never collapses to a single line.
  it('always keeps the two rails ≥ 1 device pixel apart (no collapse)', () => {
    for (const lw of [0.5, 0.75, 1, 1.1, 1.5, 2, 2.25, 3, 4.5, 6, 9]) {
      for (const dpr of [1, 2, 3]) {
        const { railDev, gapDev } = doubleRailGeometry(lw, dpr);
        expect(railDev, `railDev for lw=${lw} dpr=${dpr}`).toBeGreaterThanOrEqual(1);
        expect(gapDev, `gapDev for lw=${lw} dpr=${dpr}`).toBeGreaterThanOrEqual(1);
      }
    }
  });

  it('floors a thin double to a 1/1/1 device-pixel band', () => {
    expect(doubleRailGeometry(0.75, 1)).toEqual({ railDev: 1, gapDev: 1, spanDev: 3 });
    expect(doubleRailGeometry(1.0, 2)).toEqual({ railDev: 1, gapDev: 1, spanDev: 3 });
    expect(doubleRailGeometry(0.5, 1)).toEqual({ railDev: 1, gapDev: 1, spanDev: 3 });
  });

  it('reduces to equal thirds for a thick double', () => {
    expect(doubleRailGeometry(6, 1)).toEqual({ railDev: 2, gapDev: 2, spanDev: 6 });
    expect(doubleRailGeometry(9, 1)).toEqual({ railDev: 3, gapDev: 3, spanDev: 9 });
    expect(doubleRailGeometry(3, 2)).toEqual({ railDev: 2, gapDev: 2, spanDev: 6 });
  });

  it('keeps the band symmetric (rail === gap) and consistent', () => {
    for (const lw of [0.75, 2.25, 6]) {
      const { railDev, gapDev, spanDev } = doubleRailGeometry(lw, 2);
      expect(railDev).toBe(gapDev);
      expect(spanDev).toBe(2 * railDev + gapDev);
    }
  });
});

describe('fillDoubleBorder', () => {
  // Record fillRect calls against a minimal mock so we can assert the two rails
  // land at the right device rows/cols with a gap between them.
  function recordRects(
    fn: (ctx: { fillRect: (x: number, y: number, w: number, h: number) => void }) => void,
  ): Array<{ x: number; y: number; w: number; h: number }> {
    const rects: Array<{ x: number; y: number; w: number; h: number }> = [];
    fn({ fillRect: (x, y, w, h) => rects.push({ x, y, w, h }) });
    return rects;
  }

  it('paints a horizontal edge as two parallel rails with a gap', () => {
    // y=10, lw=0.75, dpr=2 → rail/gap = 1 device px each (span 3); band origin
    // round(10*2 − 3/2) = round(18.5) = 19 (device). Rails at device 19 and
    // 19+1+1=21, i.e. css y = 9.5 and 10.5.
    const rects = recordRects((ctx) =>
      fillDoubleBorder(ctx as never, 5, 10, 95, 10, 0.75, 2),
    );
    expect(rects).toHaveLength(2);
    // Both rails are full-width and 1 device px (0.5 css) tall.
    expect(rects[0]).toEqual({ x: 5, y: 9.5, w: 90, h: 0.5 });
    expect(rects[1]).toEqual({ x: 5, y: 10.5, w: 90, h: 0.5 });
    // A real 1-device-px gap separates the rail FILLS (rail2.top − rail1.bottom).
    expect(rects[1].y - (rects[0].y + rects[0].h)).toBeCloseTo(1 / 2);
  });

  it('paints a vertical edge as two parallel rails with a gap', () => {
    const rects = recordRects((ctx) =>
      fillDoubleBorder(ctx as never, 10, 5, 10, 95, 0.75, 2),
    );
    expect(rects).toHaveLength(2);
    expect(rects[0]).toEqual({ x: 9.5, y: 5, w: 0.5, h: 90 });
    expect(rects[1]).toEqual({ x: 10.5, y: 5, w: 0.5, h: 90 });
    expect(rects[1].x - (rects[0].x + rects[0].w)).toBeCloseTo(1 / 2);
  });
});
