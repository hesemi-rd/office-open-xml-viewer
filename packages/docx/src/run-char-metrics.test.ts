import { describe, it, expect } from 'vitest';
import {
  charScaleFactor,
  charSpacingDeltaPx,
  segAdvanceWidth,
  type LayoutTextSeg,
} from './line-layout.js';

// WD4 — pure width-model tests for the run character metrics that affect line
// breaking (§17.3.2.35 w:spacing, §17.3.2.43 w:w). These are the measure-side
// helpers that must stay measure==paint with the renderer's glyph draw. `w:kern`
// (§17.3.2.19) and `w:position` (§17.3.2.24) are exercised end-to-end by the VRT
// snapshot (they change ctx.fontKerning / the baseline y-offset, not the width
// model), so they are not unit-tested here.

function seg(partial: Partial<LayoutTextSeg>): LayoutTextSeg {
  return {
    text: 'abc',
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 12,
    color: null,
    fontFamily: null,
    vertAlign: null,
    measuredWidth: 0,
    ...partial,
  };
}

describe('WD4 run character-metric width helpers', () => {
  it('charScaleFactor defaults to 1 and reads the run w:w fraction', () => {
    expect(charScaleFactor(seg({}))).toBe(1);
    expect(charScaleFactor(seg({ charScale: 0.67 }))).toBeCloseTo(0.67, 10);
    expect(charScaleFactor(seg({ charScale: 2 }))).toBe(2);
  });

  it('charSpacingDeltaPx is the authored points scaled to px per glyph', () => {
    // 0 when the run declares no w:spacing.
    expect(charSpacingDeltaPx(seg({}), 2)).toBe(0);
    // 0.5 pt × scale 2 = 1 px per glyph.
    expect(charSpacingDeltaPx(seg({ charSpacing: 0.5 }), 2)).toBe(1);
    // Negative (tighter) tracking is preserved with sign.
    expect(charSpacingDeltaPx(seg({ charSpacing: -0.5 }), 2)).toBe(-1);
  });

  it('segAdvanceWidth scales the natural width by w:w (glyphs stretch, gaps do not)', () => {
    // No grid, no spacing: advance = natural × w:w.
    const s = seg({ text: 'abcd', charScale: 0.5 });
    expect(segAdvanceWidth(s, 100, 0, 1)).toBe(50);
  });

  it('segAdvanceWidth adds w:spacing per code point on top of the scaled width', () => {
    // "abcd" = 4 code points. natural 100 × 1.0 + 4 × (0.5 pt × scale 2 = 1 px) = 104.
    const s = seg({ text: 'abcd', charSpacing: 0.5 });
    expect(segAdvanceWidth(s, 100, 0, 2)).toBe(104);
  });

  it('segAdvanceWidth treats w:w and w:spacing independently (spacing is not stretched)', () => {
    // natural 100 × 0.8 (=80) + 4 cps × (0.5 pt × scale 1 = 0.5 px) = 82.
    const s = seg({ text: 'abcd', charScale: 0.8, charSpacing: 0.5 });
    expect(segAdvanceWidth(s, 100, 0, 1)).toBe(82);
  });

  it('segAdvanceWidth reduces to the natural width when neither attribute is set', () => {
    const s = seg({ text: 'hello' });
    expect(segAdvanceWidth(s, 123.4, 0, 1)).toBe(123.4);
  });

  it('counts surrogate-pair code points once for w:spacing', () => {
    // A single astral code point (👍) is ONE glyph, so one spacing gap.
    const s = seg({ text: '\u{1F44D}', charSpacing: 1 });
    // natural 20 + 1 cp × (1 pt × scale 1) = 21.
    expect(segAdvanceWidth(s, 20, 0, 1)).toBe(21);
  });
});
