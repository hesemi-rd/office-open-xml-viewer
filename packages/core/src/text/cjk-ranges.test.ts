import { describe, it, expect } from 'vitest';
import { isCjkBreakChar } from './cjk-ranges.js';

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
