import { describe, expect, it } from 'vitest';
import { SYMBOL_FONT_MAP, symbolFontToUnicode } from './symbol-font';

const ch = (code: number): string => String.fromCharCode(code);

describe('symbolFontToUnicode', () => {
  // The load-bearing case: list bullets. Word serializes a Symbol bullet as the
  // PUA code point U+F0B7 and a Wingdings square as U+F0A7; both the bare and
  // PUA-shifted forms must normalize so the marker draws in any fallback face.
  it('maps Symbol / Wingdings bullet code points (bare + PUA) to Unicode', () => {
    expect(symbolFontToUnicode(ch(0xf0b7), 'Symbol')).toBe('•');
    expect(symbolFontToUnicode(ch(0xb7), 'Symbol')).toBe('•');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings')).toBe('▪');
  });

  it('maps the Wingdings barb arrow block', () => {
    expect(symbolFontToUnicode(ch(0xf0e0), 'Wingdings')).toBe('→');
    expect(symbolFontToUnicode(ch(0xdf), 'Wingdings')).toBe('←');
  });

  // The gate is keyed on the requested font family, not the character — a normal
  // body font must never be remapped (a real "·" / "§" / "A" stays itself).
  it('passes characters through for non-symbol fonts', () => {
    expect(symbolFontToUnicode(ch(0xb7), 'Calibri')).toBe(ch(0xb7));
    expect(symbolFontToUnicode(ch(0xa7), 'Times New Roman')).toBe(ch(0xa7));
    expect(symbolFontToUnicode('A', 'Symbol')).toBe('A');
  });

  it('passes through when the family is missing or unmapped', () => {
    expect(symbolFontToUnicode(ch(0xf0b7), null)).toBe(ch(0xf0b7));
    expect(symbolFontToUnicode(ch(0xf0b7), undefined)).toBe(ch(0xf0b7));
    // A code point with no entry falls through unchanged rather than guessing.
    expect(symbolFontToUnicode(ch(0xfe), 'Wingdings')).toBe(ch(0xfe));
  });

  it('matches "wingdings" case-insensitively and as a substring', () => {
    expect(symbolFontToUnicode(ch(0xf0a7), 'WINGDINGS')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings 2')).toBe('▪');
  });

  // Guard against the ambiguous Symbol-vs-Wingdings code points that were
  // removed (0x24/0x4B/0x76/0xB8/0xB9/0xFE): they must NOT be in the table.
  it('omits ambiguous font-specific code points', () => {
    for (const code of [0x24, 0x4b, 0x76, 0xb8, 0xb9, 0xfe]) {
      expect(SYMBOL_FONT_MAP[code]).toBeUndefined();
    }
  });
});
