import { describe, it, expect } from 'vitest';
import { isCjkBreakChar, isLatinWordCodePoint } from './cjk-ranges.js';

describe('isCjkBreakChar', () => {
  // [code point, expected] — exercises every range edge (low−1, low, high,
  // high+1) plus the unassigned/Jamo-Extended-B gap above Hangul Syllables.
  const cases: [number, boolean][] = [
    [0x2fff, false], // just below U+3000
    [0x3000, true], // ideographic space (range start)
    [0x3001, true], // 、 (former pptx-justify start)
    [0x9fff, true], // CJK Unified Ideographs end
    [0xa000, false], // just above U+9FFF
    [0xac00, true], // Hangul Syllables start
    [0xd7a3, true], // Hangul Syllables end (canonical upper bound)
    [0xd7a4, false], // unassigned — was wrongly included by the D7AF/D7FF copies
    [0xd7ff, false], // Hangul Jamo Extended-B — was wrongly included by the D7FF copy
    [0xf8ff, false], // Private Use Area — just below CJK Compatibility Ideographs
    [0xf900, true], // CJK Compatibility Ideographs start
    [0xfaff, true], // CJK Compatibility Ideographs end
    [0xfb00, false], // just above U+FAFF (Alphabetic Presentation Forms)
    [0xfeff, false], // BOM / ZWNBSP — below U+FF00
    [0xff00, true], // Halfwidth/Fullwidth Forms start
    [0xffef, true], // Halfwidth/Fullwidth Forms end
    [0xfff0, false], // just above U+FFEF
  ];

  for (const [cp, expected] of cases) {
    it(`U+${cp.toString(16).toUpperCase().padStart(4, '0')} → ${expected}`, () => {
      expect(isCjkBreakChar(cp)).toBe(expected);
    });
  }
});

describe('isLatinWordCodePoint', () => {
  // Word characters (no intra-word break) — letters, digits, ASCII punctuation,
  // and NBSP (U+00A0, non-breaking, so it must NOT count as a break boundary).
  const wordCps: [string, number][] = [
    ['a', 0x61], ['Z', 0x5a], ['7', 0x37], ['comma', 0x2c], ['period', 0x2e],
    ['paren', 0x29], ['hyphen', 0x2d], ['at', 0x40], ['NBSP', 0x00a0],
  ];
  for (const [label, c] of wordCps) {
    it(`${label} is a word code point`, () => expect(isLatinWordCodePoint(c)).toBe(true));
  }
  // Break-eligible whitespace and CJK code points are NOT word characters.
  const nonWordCps: [string, number][] = [
    ['space', 0x20], ['tab', 0x09], ['LF', 0x0a], ['CR', 0x0d],
    ['ideograph', 0x6f22], ['ideographic comma', 0x3001], ['fullwidth comma', 0xff0c],
  ];
  for (const [label, c] of nonWordCps) {
    it(`${label} is not a word code point`, () => expect(isLatinWordCodePoint(c)).toBe(false));
  }
});
