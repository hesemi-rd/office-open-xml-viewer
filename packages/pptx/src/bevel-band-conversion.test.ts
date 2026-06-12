import { describe, it, expect } from 'vitest';
import { buildBevelInputs } from './renderer';

/**
 * Pins the EMU → device-px bevel band-width conversion (ECMA-376 §20.1.5.3: the
 * bevel `w`/`h` are EMU lengths). The slide-6 "flat cut" investigation suspected a
 * double-applied scale blowing the band up by orders of magnitude; it is NOT — the
 * band is exactly `w(EMU) × (EMU→CSS-px scale) × devScale`. These tests fail if any
 * future change re-introduces a stray scale factor, parametrised over devScale
 * {1,4,8} (the renders where the artifact was reported).
 */
describe('bevel band-width EMU→device-px conversion (no double-apply)', () => {
  // slide-6 declared values.
  const SLIDE_W_EMU = 12192000;
  const TARGET_WIDTH = 1920;
  const cssScale = TARGET_WIDTH / SLIDE_W_EMU; // EMU → CSS px
  const BEVEL_W_EMU = 304800; // 24 pt
  const BEVEL_H_EMU = 152400; // 12 pt

  const sp3d = { bevelT: { w: BEVEL_W_EMU, h: BEVEL_H_EMU, prst: 'hardEdge' } };

  for (const devScale of [1, 4, 8]) {
    it(`widthPx == w(EMU) × scale × devScale exactly [devScale ${devScale}]`, () => {
      const [bevel] = buildBevelInputs(sp3d, undefined, 'matte', cssScale, devScale);
      const expectWidth = BEVEL_W_EMU * cssScale * devScale;
      const expectHeight = BEVEL_H_EMU * cssScale * devScale;
      // ±1 px tolerance per the brief (bandPx == declared × devScale ± 1px).
      expect(Math.abs(bevel.widthPx - expectWidth)).toBeLessThanOrEqual(1);
      expect(Math.abs(bevel.heightPx - expectHeight)).toBeLessThanOrEqual(1);
      // Sanity: the band scales LINEARLY with devScale (no super-linear blow-up).
      // 24 pt at 1920 px wide ⇒ exactly 24 device px per devScale unit.
      expect(bevel.widthPx / devScale).toBeCloseTo(BEVEL_W_EMU * cssScale, 6);
    });
  }

  it('omits the bevel when w or h is zero (no spurious band)', () => {
    expect(buildBevelInputs({ bevelT: { w: 0, h: 100, prst: 'circle' } }, undefined, 'matte', cssScale, 4)).toHaveLength(0);
    expect(buildBevelInputs({ bevelT: { w: 100, h: 0, prst: 'circle' } }, undefined, 'matte', cssScale, 4)).toHaveLength(0);
  });
});
