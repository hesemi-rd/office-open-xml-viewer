import { describe, it, expect } from 'vitest';
import type { Stroke } from '@silurus/ooxml-core';
import { resolveTableBorderConflict } from './table-border-conflict.js';

// DrawingML `<a:tbl>` adjacent-cell border conflict — pure kernel. The spec is
// SILENT (no §17.4.66 analog for PresentationML), so these assert OUR defined
// deterministic rule: null loses → wider wins → darker wins → owner (a) wins.

const ln = (over: Partial<Stroke> = {}): Stroke => ({
  color: '000000',
  width: 12700, // 1pt in EMU
  ...over,
});

describe('resolveTableBorderConflict (DrawingML shared-edge, spec-silent)', () => {
  it('both null ⇒ null (no line)', () => {
    expect(resolveTableBorderConflict(null, null)).toBeNull();
  });

  it('rule #0 — a null side loses; the other real line is displayed', () => {
    const real = ln({ color: '112233' });
    expect(resolveTableBorderConflict(null, real)).toBe(real);
    expect(resolveTableBorderConflict(real, null)).toBe(real);
  });

  it('rule #1 — the wider line wins (EMU width), regardless of side', () => {
    const thin = ln({ width: 12700, color: 'ff0000' });
    const thick = ln({ width: 38100, color: '0000ff' });
    expect(resolveTableBorderConflict(thin, thick)).toBe(thick);
    expect(resolveTableBorderConflict(thick, thin)).toBe(thick);
  });

  it('rule #2 — equal width ⇒ the darker colour wins', () => {
    const dark = ln({ width: 12700, color: '000000' });
    const light = ln({ width: 12700, color: 'ffffff' });
    expect(resolveTableBorderConflict(light, dark)).toBe(dark);
    expect(resolveTableBorderConflict(dark, light)).toBe(dark);
  });

  it('rule #3 — fully tied ⇒ the owning side (a) wins', () => {
    const a = ln({ width: 12700, color: '336699' });
    const b = ln({ width: 12700, color: '336699' });
    expect(resolveTableBorderConflict(a, b)).toBe(a);
  });

  it('an 8-hex colour ignores its alpha pair for the luminance compare', () => {
    // 00000080 (translucent black) vs ffffffff (opaque white) — black is darker.
    const dark = ln({ width: 12700, color: '00000080' });
    const light = ln({ width: 12700, color: 'ffffffff' });
    expect(resolveTableBorderConflict(light, dark)).toBe(dark);
  });
});
