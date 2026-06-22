import { describe, expect, it } from 'vitest';
import {
  SYMBOL_FONT_MAP,
  SYMBOL_MAP,
  WINGDINGS_MAP,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
} from './symbol-font';

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

  // The core reason for splitting into two tables: the SAME code point resolves
  // to a DIFFERENT glyph depending on the font, because Symbol and Wingdings use
  // different encodings. A single shared table cannot be correct for both.
  it('resolves the same code point differently for Symbol vs Wingdings', () => {
    // 0xA7: Adobe Symbol → U+2663 club; Wingdings cmap → U+25AA small square.
    expect(symbolFontToUnicode(ch(0xa7), 'Symbol')).toBe('♣');
    expect(symbolFontToUnicode(ch(0xa7), 'Wingdings')).toBe('▪');
    // 0xB7: Adobe Symbol → U+2022 bullet; Wingdings cmap → a clock face
    // (astral U+1F550), which we leave as passthrough rather than mismap.
    expect(symbolFontToUnicode(ch(0xb7), 'Symbol')).toBe('•');
    expect(symbolFontToUnicode(ch(0xb7), 'Wingdings')).toBe(ch(0xb7));
    // 0xB8: Adobe Symbol → U+00F7 division; Wingdings → a clock face (passthrough).
    expect(symbolFontToUnicode(ch(0xb8), 'Symbol')).toBe('÷');
    expect(symbolFontToUnicode(ch(0xb8), 'Wingdings')).toBe(ch(0xb8));
  });

  // Restored ambiguous entries (removed in PR #519) now carry the correct
  // font-specific value rather than a single guessed glyph.
  it('restores the previously-removed code points with font-specific values', () => {
    // Symbol side
    expect(symbolFontToUnicode(ch(0xb8), 'Symbol')).toBe('÷'); // U+00F7
    expect(symbolFontToUnicode(ch(0xb9), 'Symbol')).toBe('≠'); // U+2260
    // Wingdings side
    expect(symbolFontToUnicode(ch(0x24), 'Wingdings')).toBe('\u{1F453}'); // eyeglasses
    expect(symbolFontToUnicode(ch(0x4b), 'Wingdings')).toBe('\u{1F610}'); // neutral face
    expect(symbolFontToUnicode(ch(0x76), 'Wingdings')).toBe('❖'); // U+2756
    expect(symbolFontToUnicode(ch(0xfe), 'Wingdings')).toBe('☑'); // U+2611 ballot box w/ check
  });

  it('maps the Wingdings barb arrow block', () => {
    expect(symbolFontToUnicode(ch(0xf0e0), 'Wingdings')).toBe('→');
    expect(symbolFontToUnicode(ch(0xdf), 'Wingdings')).toBe('←');
  });

  // Geometric markers. 0x74 lozenge6 / 0x77 rhombus4 are both SOLID glyphs —
  // both must be the BLACK diamond, never the white ◇ (regression guard for the
  // 0x77 fill error).
  it('maps the Wingdings solid geometric markers to black glyphs', () => {
    expect(symbolFontToUnicode(ch(0x6c), 'Wingdings')).toBe('●');
    expect(symbolFontToUnicode(ch(0x6e), 'Wingdings')).toBe('■');
    expect(symbolFontToUnicode(ch(0x74), 'Wingdings')).toBe('◆');
    expect(symbolFontToUnicode(ch(0x77), 'Wingdings')).toBe('◆');
  });

  it('maps Symbol arrows from the Adobe encoding', () => {
    // symbol.txt: 0xAE → U+2192 RIGHTWARDS ARROW, 0xAC → U+2190 LEFTWARDS ARROW
    expect(symbolFontToUnicode(ch(0xae), 'Symbol')).toBe('→');
    expect(symbolFontToUnicode(ch(0xac), 'Symbol')).toBe('←');
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
    // A code point with no entry in the selected table falls through unchanged
    // rather than guessing (e.g. 0x21 is "!" in Symbol, not a dingbat).
    expect(symbolFontToUnicode(ch(0x21), 'Symbol')).toBe(ch(0x21));
  });

  it('matches "wingdings" case-insensitively and as a substring', () => {
    expect(symbolFontToUnicode(ch(0xf0a7), 'WINGDINGS')).toBe('▪');
    expect(symbolFontToUnicode(ch(0xf0a7), 'Wingdings 2')).toBe('▪');
  });

  // Pencil/scissors/faces are Wingdings glyphs only. In Symbol, 0x21 is "!" and
  // 0x22 is "∀" — so these dingbats must NOT leak across the font gate.
  it('keeps Wingdings dingbats out of the Symbol table', () => {
    expect(SYMBOL_MAP[0x21]).toBeUndefined();
    expect(SYMBOL_MAP[0x24]).toBeUndefined();
    expect(symbolFontToUnicode(ch(0x24), 'Symbol')).toBe(ch(0x24));
  });

  // The deprecated shared alias still resolves Wingdings markers (it aliases the
  // Wingdings table) so any legacy importer keeps working.
  it('keeps the deprecated SYMBOL_FONT_MAP alias pointing at Wingdings', () => {
    expect(SYMBOL_FONT_MAP).toBe(WINGDINGS_MAP);
    expect(SYMBOL_FONT_MAP[0xa7]).toBe('▪');
  });
});

describe('isSymbolFontFamily', () => {
  it('matches exactly "symbol" and any "wingdings" family', () => {
    expect(isSymbolFontFamily('Symbol')).toBe(true);
    expect(isSymbolFontFamily('SYMBOL')).toBe(true);
    expect(isSymbolFontFamily('Wingdings')).toBe(true);
    expect(isSymbolFontFamily('Wingdings 2')).toBe(true);
  });

  it('does NOT match Webdings (no core table) or ordinary faces', () => {
    // Webdings is intentionally excluded — core has no Webdings table, so the
    // gate must not claim it can normalize a Webdings run.
    expect(isSymbolFontFamily('Webdings')).toBe(false);
    expect(isSymbolFontFamily('Calibri')).toBe(false);
    // "Symbol" must be exact, not a substring (e.g. "SymbolMT" is a real Latin
    // PostScript face), matching symbolFontToUnicode's own gate.
    expect(isSymbolFontFamily('SymbolMT')).toBe(false);
    expect(isSymbolFontFamily(null)).toBe(false);
    expect(isSymbolFontFamily(undefined)).toBe(false);
  });
});

describe('symbolTextToUnicodeSegments', () => {
  it('returns the whole string unmapped for a non-symbol font', () => {
    expect(symbolTextToUnicodeSegments('abc', 'Calibri')).toEqual([
      { text: 'abc', mapped: false },
    ]);
    expect(symbolTextToUnicodeSegments(ch(0xf0a7), null)).toEqual([
      { text: ch(0xf0a7), mapped: false },
    ]);
  });

  it('maps every Wingdings PUA glyph and flags the run as mapped', () => {
    // U+F0A7 (square) U+F0E0 (right arrow) → ▪ →
    const segs = symbolTextToUnicodeSegments(ch(0xf0a7) + ch(0xf0e0), 'Wingdings');
    expect(segs).toEqual([{ text: '▪→', mapped: true }]);
  });

  it('splits at mapped/unmapped boundaries so each run is single-font', () => {
    // 0x21 "!" has no Wingdings entry (passthrough), 0xF0A7 maps to ▪.
    const segs = symbolTextToUnicodeSegments(ch(0x21) + ch(0xf0a7) + ch(0x22), 'Wingdings');
    expect(segs).toEqual([
      { text: ch(0x21), mapped: false },
      { text: '▪', mapped: true },
      { text: ch(0x22), mapped: false },
    ]);
  });

  it('preserves astral targets by iterating code points', () => {
    // Wingdings 0x24 → U+1F453 EYEGLASSES (astral). Must survive as one glyph.
    const segs = symbolTextToUnicodeSegments(ch(0x24), 'Wingdings');
    expect(segs).toEqual([{ text: '\u{1F453}', mapped: true }]);
  });

  it('handles the empty string', () => {
    expect(symbolTextToUnicodeSegments('', 'Wingdings')).toEqual([
      { text: '', mapped: false },
    ]);
  });
});
