import { describe, it, expect } from 'vitest';
import { justifyLine } from './text-justify.js';

// A laid-out line is usually ONE merged segment per style run, so the tests
// feed whole strings (not pre-split tokens) — that is the real input shape.
type Seg = { text?: string; tag?: string };

/**
 * Collapse each justified segment to a compact tuple for assertions:
 * `[text, jext, splitBefore?, perGap?]`. splitBefore and perGap are only
 * included when the segment has internal CJK gaps.
 */
const annot = (r: ReturnType<typeof justifyLine<Seg>>) =>
  r === null
    ? null
    : r.map((s) => {
        const j = +s.jext.toFixed(3);
        return s.splitBefore && s.splitBefore.length > 0
          ? ([s.text, j, s.splitBefore, +(s.perGap ?? 0).toFixed(3)] as const)
          : ([s.text, j] as const);
      });

describe('justifyLine', () => {
  it("returns null for 'just' on the last line (short sentences stay natural)", () => {
    expect(justifyLine<Seg>([{ text: '日本語' }], 120, 60, 'just', true)).toBeNull();
  });

  it("'dist' justifies even the last line", () => {
    // 3 CJK glyphs → 2 internal gaps in this single segment; slack 60 → perGap 30.
    const r = justifyLine<Seg>([{ text: '日本語' }], 120, 60, 'dist', true);
    expect(annot(r)).toEqual([['日本語', 0, [1, 2], 30]]);
  });

  it('pure Latin: widens inter-word spaces (kernel splits before the space-following word)', () => {
    // "Hello world foo": gaps fall before 'w' (code-point 6) and before 'f' (12).
    // slack 100, 2 gaps → perGap 50, no trailing gap.
    const r = justifyLine<Seg>([{ text: 'Hello world foo' }], 200, 100, 'just', false);
    expect(annot(r)).toEqual([['Hello world foo', 0, [6, 12], 50]]);
  });

  it('pure CJK: widens every inter-character gap except after the final glyph', () => {
    const r = justifyLine<Seg>([{ text: '日本語' }], 120, 60, 'just', false);
    expect(annot(r)).toEqual([['日本語', 0, [1, 2], 30]]);
  });

  it('mixed EC市場で: no gap inside the Latin "EC", gaps at C|市, 市|場, 場|で', () => {
    // 5 code points, 3 gaps inside (after positions 2, 3, 4). slack 30 → perGap 10.
    const r = justifyLine<Seg>([{ text: 'EC市場で' }], 130, 100, 'just', false);
    expect(annot(r)).toEqual([['EC市場で', 0, [2, 3, 4], 10]]);
  });

  it('leading 字下げ whitespace stays fixed (no stretch before first content)', () => {
    // "  日本語": gap after first content boundary (before code-point 3 = 本) and (before 4 = 語).
    // The leading "  " contributes no gap. slack 30 → perGap 15.
    const r = justifyLine<Seg>([{ text: '  日本語' }], 130, 100, 'just', false);
    expect(annot(r)).toEqual([['  日本語', 0, [3, 4], 15]]);
  });

  it('trailing whitespace at line end does not stretch', () => {
    // "日本 ": one gap before 本 (cp 1). slack 40 → perGap 40, no trailing.
    const r = justifyLine<Seg>([{ text: '日本 ' }], 100, 60, 'just', false);
    expect(annot(r)).toEqual([['日本 ', 0, [1], 40]]);
  });

  it('evaluates CJK boundaries across a style (segment) boundary', () => {
    // [日本][語]: gaps after 日 (inside seg 1) and after 本 (between seg 1 and seg 2 → trailing).
    // slack 60, 2 gaps → perGap 30.
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
    expect(r!.map((s) => [s.text, +s.jext.toFixed(3), s.splitBefore ?? null, s.tag])).toEqual([
      ['日本', 30, [1], 'a'],
      ['語', 0, null, 'b'],
    ]);
  });

  it('opens gaps INSIDE the final segment too (CJK boundary inside the last run)', () => {
    // [日本][語学], 4 glyphs. PowerPoint widens every inter-CJK boundary except
    // after the last glyph — INCLUDING 語|学, which lives inside the final
    // segment. Only the final glyph's gap is suppressed (the content-span trim),
    // not the whole final segment. Pins that the adapter excludes no segment by
    // index (lastDrawnSi sentinel = segments.length).
    const r = justifyLine<Seg>(
      [
        { text: '日本', tag: 'a' },
        { text: '語学', tag: 'b' },
      ],
      150,
      60,
      'just',
      false,
    );
    expect(r).not.toBeNull();
    expect(r!.map((s) => [s.text, +s.jext.toFixed(3), s.splitBefore ?? null, s.tag])).toEqual([
      ['日本', 30, [1], 'a'],
      ['語学', 0, [1], 'b'],
    ]);
  });

  it('an inline object (text===undefined) is one unit and can take a gap', () => {
    // gaps after 日 and after the object → perGap = 40/2 = 20
    const r = justifyLine<Seg>([{ text: '日' }, {}, { text: '本' }], 100, 60, 'just', false);
    expect(r).not.toBeNull();
    expect(r!.map((s) => [s.text, +s.jext.toFixed(3)])).toEqual([
      ['日', 20],
      [undefined, 20],
      ['本', 0],
    ]);
  });

  it('emits an empty-text segment (inline OMML equation) so it is still drawn', () => {
    const r = justifyLine<Seg>(
      [
        { text: '日', tag: 'a' },
        { text: '', tag: 'math' },
        { text: '本', tag: 'b' },
      ],
      120,
      60,
      'just',
      false,
    );
    expect(r).not.toBeNull();
    expect(r!.map((s) => [s.text, +s.jext.toFixed(3), s.tag])).toEqual([
      ['日', 60, 'a'],
      ['', 0, 'math'],
      ['本', 0, 'b'],
    ]);
  });

  it('Σ jext + Σ (perGap × splits) equals slack (line reaches availWidth)', () => {
    const r = justifyLine<Seg>([{ text: 'Hello world foo bar' }], 300, 120, 'just', false);
    expect(r).not.toBeNull();
    const sum = r!.reduce((acc, s) => {
      const internal = s.splitBefore && s.perGap ? s.splitBefore.length * s.perGap : 0;
      return acc + s.jext + internal;
    }, 0);
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

  it('preserves all style fields on every annotated segment', () => {
    const r = justifyLine<Seg>([{ text: '日本語', tag: 'x' }], 120, 60, 'just', false);
    expect(r!.every((s) => s.tag === 'x')).toBe(true);
  });

  // Regression for the CJK-range dedup (isCjkBreakChar includes U+3000).
  // U+3000 (ideographic space) is whitespace per /\s/, so it is stretched as a
  // single inter-word gap and never reaches the CJK predicate.
  it('U+3000 stretches as one inter-word gap, not double-counted as a CJK gap', () => {
    // "日　本": one gap before code-point 2 (本). slack 60 → perGap 60.
    const r = justifyLine<Seg>([{ text: '日　本' }], 120, 60, 'just', false);
    expect(annot(r)).toEqual([['日　本', 0, [2], 60]]);
  });

  it('U+3000 between CJK glyphs: single stretch point even with neighbours', () => {
    // 日　本で: 4 code points, gaps before 本 (cp 2) and before で (cp 3). slack 80 → perGap 40.
    const r = justifyLine<Seg>([{ text: '日　本で' }], 140, 60, 'just', false);
    expect(annot(r)).toEqual([['日　本で', 0, [2, 3], 40]]);
  });
});
