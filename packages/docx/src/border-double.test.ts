import { describe, expect, it } from 'vitest';
import { doubleRailGeometry } from './renderer.js';

// ECMA-376 §17.18.2 ST_Border "double": two parallel lines. The standard leaves
// the rail/gap PIXEL geometry to the implementation; we render three device-px
// bands (rail / gap / rail, each ≈ lw/3) with each band floored at one device
// pixel so a thin double never collapses into a single line. `lw` is the stroked
// width in px; the canvas is `ctx.scale(dpr,dpr)`-d, so the band is laid out in
// device pixels (lw·dpr).
describe('doubleRailGeometry (ST_Border §17.18.2 "double")', () => {
  // THE regression gate: a thin `double` (e.g. a table "Total" row top at
  // sz6 ≈ 0.75px) used to render the two rails overlapping with no gap, so it
  // looked like a single line. For ANY width/dpr the two rails must stay
  // separated — gapDev ≥ 1 device px — and each rail must paint — railDev ≥ 1.
  it('always keeps the two rails ≥ 1 device pixel apart (no collapse)', () => {
    for (const lw of [0.5, 0.75, 1, 1.1, 1.5, 2, 2.25, 3, 4.5, 6, 9]) {
      for (const dpr of [1, 2, 3]) {
        const { railDev, gapDev } = doubleRailGeometry(lw, dpr);
        expect(railDev, `railDev for lw=${lw} dpr=${dpr}`).toBeGreaterThanOrEqual(1);
        expect(gapDev, `gapDev for lw=${lw} dpr=${dpr}`).toBeGreaterThanOrEqual(1);
      }
    }
  });

  // A thin double (sz6 = 0.75pt) at scale≈1: the case that used to collapse.
  // The one-device-pixel floor lifts the 0-rounded thirds to a 1/1/1 band.
  it('floors a thin double to a 1/1/1 device-pixel band', () => {
    expect(doubleRailGeometry(0.75, 1)).toEqual({ railDev: 1, gapDev: 1, spanDev: 3 });
    expect(doubleRailGeometry(1.0, 2)).toEqual({ railDev: 1, gapDev: 1, spanDev: 3 });
    // Even a sub-floor hairline still yields a visible two-line band.
    expect(doubleRailGeometry(0.5, 1)).toEqual({ railDev: 1, gapDev: 1, spanDev: 3 });
  });

  // A thick double reduces to the equal line/gap/line thirds (no floor effect).
  it('reduces to equal thirds for a thick double', () => {
    expect(doubleRailGeometry(6, 1)).toEqual({ railDev: 2, gapDev: 2, spanDev: 6 });
    expect(doubleRailGeometry(9, 1)).toEqual({ railDev: 3, gapDev: 3, spanDev: 9 });
    expect(doubleRailGeometry(3, 2)).toEqual({ railDev: 2, gapDev: 2, spanDev: 6 });
  });

  // The band is symmetric — rail and gap are the same width at every size — and
  // spanDev is always 2·railDev + gapDev.
  it('keeps the band symmetric (rail === gap) and consistent', () => {
    for (const lw of [0.75, 2.25, 6]) {
      const { railDev, gapDev, spanDev } = doubleRailGeometry(lw, 2);
      expect(railDev).toBe(gapDev);
      expect(spanDev).toBe(2 * railDev + gapDev);
    }
  });
});
