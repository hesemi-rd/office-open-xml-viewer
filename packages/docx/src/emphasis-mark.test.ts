import { describe, it, expect } from 'vitest';
import { emphasisMarkCenters, emphasisMarkGeometry } from './emphasis-mark.js';

// ECMA-376 §17.3.2.12 w:em / §17.18.24 ST_Em — pure geometry of the emphasis
// (圏点) marks. These functions decide WHERE each per-glyph mark goes and WHAT
// shape it is, independent of the canvas; the renderer stamps them.

// A measure that gives every code point a fixed 10px advance, so a mark's centre
// under glyph i must land at penX + (i + 0.5) * 10 (+ pitch terms).
const uniform10 = (s: string): number => [...s].length * 10;

describe('emphasisMarkCenters (§17.18.24 — one mark per non-space character)', () => {
  it('centres a mark under each glyph at the midpoint of its advance', () => {
    const c = emphasisMarkCenters('abc', uniform10, 0, 0);
    expect(c.map((p) => p.centerX)).toEqual([5, 15, 25]);
  });

  it('offsets every centre by the pen x', () => {
    const c = emphasisMarkCenters('ab', uniform10, 100, 0);
    expect(c.map((p) => p.centerX)).toEqual([105, 115]);
  });

  it('skips space characters but still advances the cumulative measure', () => {
    // "a b" — the middle space gets NO mark, but 'b' still lands at its 3rd cell.
    const c = emphasisMarkCenters('a b', uniform10, 0, 0);
    expect(c.map((p) => p.centerX)).toEqual([5, 25]);
  });

  it('skips the ideographic space U+3000 too (whitespace class)', () => {
    const c = emphasisMarkCenters('あ　い', uniform10, 0, 0);
    // Two marks (あ, い); the wide space at index 1 is skipped.
    expect(c).toHaveLength(2);
    expect(c.map((p) => p.centerX)).toEqual([5, 25]);
  });

  it('adds the uniform per-glyph pitch between glyphs (docGrid / justify)', () => {
    // pitch = 4px: glyph i occupies [i*10 + i*4, (i+1)*10 + (i+1)*4], centre =
    // ((i*10 + i*4) + ((i+1)*10 + (i+1)*4)) / 2 = i*14 + 7.
    const c = emphasisMarkCenters('abc', uniform10, 0, 4);
    expect(c.map((p) => p.centerX)).toEqual([7, 21, 35]);
  });

  it('keeps surrogate-pair code points as a single glyph', () => {
    // A single astral code point (U+20B9F 𠮟) is ONE mark, not two UTF-16 units.
    const c = emphasisMarkCenters('\u{20B9F}', uniform10, 0, 0);
    expect(c).toHaveLength(1);
    expect(c[0].centerX).toBe(5);
  });
});

describe('emphasisMarkGeometry (§17.18.24 ST_Em value → shape/position)', () => {
  it('dot → filled disc above the glyphs', () => {
    const g = emphasisMarkGeometry('dot', 40);
    expect(g.shape).toBe('dot');
    expect(g.above).toBe(true);
    expect(g.radius).toBeCloseTo(40 * 0.07, 6);
  });

  it('circle → hollow circle above the glyphs', () => {
    const g = emphasisMarkGeometry('circle', 40);
    expect(g.shape).toBe('circle');
    expect(g.above).toBe(true);
  });

  it('comma → sesame/comma shape above the glyphs', () => {
    const g = emphasisMarkGeometry('comma', 40);
    expect(g.shape).toBe('comma');
    expect(g.above).toBe(true);
  });

  it('underDot → filled disc BELOW the glyphs (§17.18.24)', () => {
    const g = emphasisMarkGeometry('underDot', 40);
    expect(g.shape).toBe('dot');
    expect(g.above).toBe(false);
  });

  it('mark radius scales with the effective font size', () => {
    expect(emphasisMarkGeometry('dot', 20).radius).toBeCloseTo(
      emphasisMarkGeometry('dot', 40).radius / 2,
      6,
    );
  });
});
