import { describe, it, expect } from 'vitest';
import { justifyLine } from './text-justify.js';

// A laid-out line is usually ONE merged segment per style run, so the tests
// feed whole strings (not pre-split tokens) — that is the real input shape.
type Seg = { text?: string; tag?: string };

// Collapse pieces to [text, roundedJext] tuples for compact assertions.
const pieces = (r: ReturnType<typeof justifyLine<Seg>>) =>
  r === null ? null : r.map((p) => [p.text, +p.jext.toFixed(3)] as const);

describe('justifyLine', () => {
  it("returns null for 'just' on the last line (short sentences stay natural)", () => {
    expect(justifyLine<Seg>([{ text: '日本語' }], 120, 60, 'just', true)).toBeNull();
  });

  it("'dist' justifies even the last line", () => {
    const r = justifyLine<Seg>([{ text: '日本語' }], 120, 60, 'dist', true);
    expect(pieces(r)).toEqual([
      ['日', 30],
      ['本', 30],
      ['語', 0],
    ]);
  });

  it('pure Latin: widens each inter-word space, splitting after the spaces', () => {
    const r = justifyLine<Seg>([{ text: 'Hello world foo' }], 200, 100, 'just', false);
    expect(pieces(r)).toEqual([
      ['Hello ', 50],
      ['world ', 50],
      ['foo', 0],
    ]);
  });

  it('pure CJK: widens every inter-character gap except after the final glyph', () => {
    const r = justifyLine<Seg>([{ text: '日本語' }], 120, 60, 'just', false);
    expect(pieces(r)).toEqual([
      ['日', 30],
      ['本', 30],
      ['語', 0],
    ]);
  });

  it('mixed EC市場で: no gap inside the Latin "EC", gaps at C|市, 市|場, 場|で', () => {
    const r = justifyLine<Seg>([{ text: 'EC市場で' }], 130, 100, 'just', false);
    expect(pieces(r)).toEqual([
      ['EC', 10],
      ['市', 10],
      ['場', 10],
      ['で', 0],
    ]);
  });

  it('leading 字下げ whitespace stays fixed (no stretch before first content)', () => {
    const r = justifyLine<Seg>([{ text: '  日本語' }], 130, 100, 'just', false);
    expect(pieces(r)).toEqual([
      ['  日', 15],
      ['本', 15],
      ['語', 0],
    ]);
  });

  it('trailing whitespace at line end does not stretch', () => {
    const r = justifyLine<Seg>([{ text: '日本 ' }], 100, 60, 'just', false);
    expect(pieces(r)).toEqual([
      ['日', 40],
      ['本 ', 0],
    ]);
  });

  it('evaluates CJK boundaries across a style (segment) boundary', () => {
    const r = justifyLine<Seg>(
      [
        { text: '日本', tag: 'a' },
        { text: '語', tag: 'b' },
      ],
      120,
      60,
      'just',
      false,
    );
    expect(r).not.toBeNull();
    expect(r!.map((p) => [p.text, +p.jext.toFixed(3), p.tag])).toEqual([
      ['日', 30, 'a'],
      ['本', 30, 'a'],
      ['語', 0, 'b'],
    ]);
  });

  it('an inline object (text===undefined) is one unit and can take a gap', () => {
    const r = justifyLine<Seg>([{ text: '日' }, {}, { text: '本' }], 100, 60, 'just', false);
    // gaps after 日 and after the object → perGap = 40/2 = 20
    expect(r).not.toBeNull();
    expect(r!.map((p) => [p.text, +p.jext.toFixed(3)])).toEqual([
      ['日', 20],
      [undefined, 20],
      ['本', 0],
    ]);
  });

  it('the jext advances sum to the slack (line reaches availWidth)', () => {
    const r = justifyLine<Seg>([{ text: 'Hello world foo bar' }], 300, 120, 'just', false);
    const sum = r!.reduce((a, p) => a + p.jext, 0);
    expect(sum).toBeCloseTo(180, 6);
  });

  it('returns null when there is no slack to distribute', () => {
    expect(justifyLine<Seg>([{ text: '日本語' }], 60.2, 60, 'just', false)).toBeNull();
  });

  it('returns null for a single Latin word (no inner gap)', () => {
    expect(justifyLine<Seg>([{ text: 'Hello' }], 200, 50, 'just', false)).toBeNull();
  });

  it('returns null for a single CJK glyph (one content unit)', () => {
    expect(justifyLine<Seg>([{ text: '日' }], 200, 20, 'just', false)).toBeNull();
  });

  it('preserves all style fields on every split piece', () => {
    const r = justifyLine<Seg>([{ text: '日本語', tag: 'x' }], 120, 60, 'just', false);
    expect(r!.every((p) => p.tag === 'x')).toBe(true);
  });

  // Regression for the CJK-range dedup (isCjkBreakChar now includes U+3000).
  // U+3000 (ideographic space) is whitespace per /\s/, so it is stretched as a
  // single inter-word gap and never reaches the CJK predicate. Including U+3000
  // in the shared predicate must NOT add a second (inter-CJK) gap around it: the
  // 日→U+3000 boundary takes the boundary-into-whitespace path and contributes
  // no gap of its own. Output stays exactly one gap on the space.
  it('U+3000 stretches as one inter-word gap, not double-counted as a CJK gap', () => {
    const r = justifyLine<Seg>([{ text: '日　本' }], 120, 60, 'just', false);
    expect(pieces(r)).toEqual([
      ['日　', 60],
      ['本', 0],
    ]);
  });

  it('U+3000 between CJK glyphs: single stretch point even with neighbours', () => {
    // 日 本 で  → units: 日, U+3000, 本, で. Only two gaps: the U+3000 space and
    // the 本|で inter-CJK boundary (the 日|U+3000 boundary adds none). slack 80,
    // 2 gaps → perGap 40.
    const r = justifyLine<Seg>([{ text: '日　本で' }], 140, 60, 'just', false);
    expect(pieces(r)).toEqual([
      ['日　', 40],
      ['本', 40],
      ['で', 0],
    ]);
  });
});
