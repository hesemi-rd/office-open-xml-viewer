import { afterEach, describe, expect, it } from 'vitest';
import {
  containsSeaScript,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
  isSeaScriptCodePoint,
  resetSeaSegmenterForTest,
  seaWordBreakOffsets,
  setSeaWordSegmenterForTest,
} from './sea-break.js';

afterEach(() => resetSeaSegmenterForTest());

describe('isSeaScriptCodePoint', () => {
  const cases: [number, boolean][] = [
    [0x0dff, false], // just below Thai
    [0x0e00, true], // Thai start
    [0x0e33, true], // THAI SARA AM
    [0x0e50, true], // THAI DIGIT ZERO
    [0x0e7f, true], // Thai end
    [0x0e80, true], // Lao start
    [0x0eff, true], // Lao end
    [0x0f00, false], // Tibetan — out of scope
    [0x177f, false], // just below Khmer
    [0x1780, true], // Khmer start
    [0x17d2, true], // KHMER SIGN COENG
    [0x17ff, true], // Khmer end
    [0x1800, false], // just above Khmer
    [0x0041, false], // Latin A
    [0x3042, false], // Hiragana — CJK, not SEA
  ];
  for (const [cp, expected] of cases) {
    it(`U+${cp.toString(16).toUpperCase().padStart(4, '0')} → ${expected}`, () => {
      expect(isSeaScriptCodePoint(cp)).toBe(expected);
    });
  }
});

describe('containsSeaScript', () => {
  it('true for Thai / Lao / Khmer text', () => {
    expect(containsSeaScript('ภาษาไทย')).toBe(true);
    expect(containsSeaScript('ພາສາລາວ')).toBe(true);
    expect(containsSeaScript('ភាសាខ្មែរ')).toBe(true);
  });
  it('true when SEA is embedded in Latin', () => {
    expect(containsSeaScript('Hello ภาษา world')).toBe(true);
  });
  it('false for pure Latin / CJK / empty', () => {
    expect(containsSeaScript('Hello, world!')).toBe(false);
    expect(containsSeaScript('日本語のテキスト')).toBe(false);
    expect(containsSeaScript('')).toBe(false);
  });
});

describe('seaWordBreakOffsets — platform ICU', () => {
  it('returns interior grapheme-boundary offsets for a Thai sentence', () => {
    const text = 'ภาษาไทยเป็นภาษาที่สวยงามมาก';
    const offsets = seaWordBreakOffsets(text);
    expect(offsets.length).toBeGreaterThan(0);
    // Ascending, strictly interior (never 0 or text.length).
    for (let k = 1; k < offsets.length; k++) expect(offsets[k]).toBeGreaterThan(offsets[k - 1]);
    for (const o of offsets) {
      expect(o).toBeGreaterThan(0);
      expect(o).toBeLessThan(text.length);
    }
    // Every offset is a grapheme-cluster boundary (never mid-cluster).
    const graphemes = new Set(graphemeClusterOffsets(text));
    for (const o of offsets) expect(graphemes.has(o)).toBe(true);
    // Every offset falls between two SEA-script characters (interior to the span).
    for (const o of offsets) {
      expect(isSeaScriptCodePoint(text.codePointAt(o - 1)!)).toBe(true);
      expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
    }
  });

  it('dictionary-breaks Lao and Khmer too', () => {
    expect(seaWordBreakOffsets('ພາສາລາວແມ່ນພາສາທີ່ສວຍງາມ').length).toBeGreaterThan(0);
    expect(seaWordBreakOffsets('ភាសាខ្មែរជាភាសាដ៏ស្រស់ស្អាត').length).toBeGreaterThan(0);
  });

  it('never adds a break touching non-SEA text (mixed Thai + Latin + digits)', () => {
    const text = 'ทดสอบ ABC 123 ภาษาไทย';
    for (const o of seaWordBreakOffsets(text)) {
      // Both sides of every returned offset must be SEA — no Latin/space/digit edge.
      expect(isSeaScriptCodePoint(text.codePointAt(o - 1)!)).toBe(true);
      expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
    }
  });

  it('offsets are correct after an astral (surrogate-pair) non-SEA prefix (UTF-16 units)', () => {
    const emoji = '😀'; // U+1F600, 2 UTF-16 units
    const thai = 'ภาษาไทยเป็นภาษา';
    const text = emoji + thai;
    const offsets = seaWordBreakOffsets(text);
    expect(offsets.length).toBeGreaterThan(0);
    for (const o of offsets) {
      expect(o).toBeGreaterThanOrEqual(emoji.length); // never inside the emoji
      expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
    }
  });

  it('returns [] for non-SEA / empty input', () => {
    expect(seaWordBreakOffsets('Hello world')).toEqual([]);
    expect(seaWordBreakOffsets('')).toEqual([]);
  });
});

describe('seaWordBreakOffsets — graceful fallback + filtering (injected segmenter)', () => {
  it('returns [] when the segmenter is unavailable', () => {
    setSeaWordSegmenterForTest(null);
    expect(seaWordBreakOffsets('ภาษาไทยเป็นภาษา')).toEqual([]);
  });

  it('returns [] when segmentation throws', () => {
    setSeaWordSegmenterForTest(() => {
      throw new Error('boom');
    });
    expect(seaWordBreakOffsets('ภาษาไทยเป็นภาษา')).toEqual([]);
  });

  it('excludes non-word (punctuation) segment starts, keeps word starts', () => {
    // Fake: split the span into fixed segments; one is punctuation (isWordLike:false).
    // span "AAAA.BBBB" (all treated as one SEA span by the test text below).
    setSeaWordSegmenterForTest((span) => {
      // segment boundaries at 0 (word), 4 (punct '.'), 5 (word)
      return [
        { index: 0, isWordLike: true },
        { index: 4, isWordLike: false },
        { index: 5, isWordLike: true },
      ];
    });
    // Use real Thai so containsSeaScript passes; the fake decides the offsets.
    const text = 'ก'.repeat(9);
    // index 0 excluded (span-start); index 4 excluded (punctuation); index 5 kept.
    expect(seaWordBreakOffsets(text)).toEqual([5]);
  });
});

describe('fitSeaWordPrefix', () => {
  // Each char is width 10; offsets at 2,5,7 (word boundaries), text length 9.
  const text = 'abcdefghi';
  const offsets = [2, 5, 7];
  const measure = (s: string) => s.length * 10;

  it('returns the largest word boundary whose prefix fits', () => {
    // avail 55 → prefixes: [0,2)=20, [0,5)=50, [0,7)=70(overflow) → best 5.
    expect(fitSeaWordPrefix(text, offsets, 0, 55, measure)).toBe(5);
  });
  it('returns text.length when the whole remainder fits', () => {
    expect(fitSeaWordPrefix(text, offsets, 0, 999, measure)).toBe(9);
  });
  it('returns start (no progress) when even the first word overflows', () => {
    // avail 15 < first word [0,2)=20 → start.
    expect(fitSeaWordPrefix(text, offsets, 0, 15, measure)).toBe(0);
  });
  it('respects a non-zero start (continuation of a wrapped word run)', () => {
    // from start 5: boundaries >5 = [7]; avail 15 → [5,7)=20 overflow → best 5.
    expect(fitSeaWordPrefix(text, offsets, 5, 15, measure)).toBe(5);
    // avail 45 → [5,7)=20 ok, [5,9)=40 ok → 9.
    expect(fitSeaWordPrefix(text, offsets, 5, 45, measure)).toBe(9);
  });
  it('treats the final segment (past the last offset) as a whole-remainder word', () => {
    // from start 7: no offsets >7; [7,9)=20; avail 25 → 9; avail 15 → 7 (no progress).
    expect(fitSeaWordPrefix(text, offsets, 7, 25, measure)).toBe(9);
    expect(fitSeaWordPrefix(text, offsets, 7, 15, measure)).toBe(7);
  });

  it('keeps a later-fitting boundary when advance is non-monotone (negative spacing)', () => {
    // Simulate strong negative letter spacing: a longer prefix can be NARROWER
    // than a shorter one, so an early-return greedy would wrongly stop short.
    // widths: [0,2)=30, [0,5)=10, [0,7)=40, whole [0,9)=50; avail 15 → best 5.
    const w: Record<number, number> = { 2: 30, 5: 10, 7: 40, 9: 50 };
    const nonMono = (s: string) => w[s.length] ?? s.length;
    expect(fitSeaWordPrefix(text, offsets, 0, 15, nonMono)).toBe(5);
  });
});

describe('graphemeClusterOffsets', () => {
  it('keeps a base + combining/tone mark together (Thai SARA AM cluster)', () => {
    // U+0E33 SARA AM is a single code point but base+mark cluster candidates like
    // นํ (NO NU + NIKHAHIT) must not split. Use กำ = ก + SARA AM already atomic;
    // test the classic ก + tone mark ่ (U+0E48) cluster instead.
    const text = 'ก่ข้'; // ก+MAI EK, ข+MAI THO → two clusters
    const offsets = graphemeClusterOffsets(text);
    // Exactly one interior boundary, between the two clusters (after index 2).
    expect(offsets).toEqual([2]);
  });
  it('returns interior offsets only (excludes 0 and length)', () => {
    const offsets = graphemeClusterOffsets('abc');
    for (const o of offsets) {
      expect(o).toBeGreaterThan(0);
      expect(o).toBeLessThan(3);
    }
  });
});
