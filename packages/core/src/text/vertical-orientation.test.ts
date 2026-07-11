import { describe, it, expect } from 'vitest';
import {
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
} from './vertical-orientation.js';

const cp = (ch: string): number => ch.codePointAt(0) ?? 0;

describe('verticalOrientation (UAX #50 Vertical_Orientation)', () => {
  // Representative points across the four values, taken directly from
  // VerticalOrientation-17.0.0.txt.
  const cases: Array<[string, number, 'U' | 'R' | 'Tu' | 'Tr']> = [
    // U — upright: CJK ideographs, kana, Hangul, fullwidth forms.
    ['漢 CJK Unified', cp('漢'), 'U'],
    ['あ Hiragana', cp('あ'), 'U'],
    ['カ Katakana', cp('カ'), 'U'],
    ['가 Hangul syllable', cp('가'), 'U'],
    ['Ａ fullwidth Latin A (FF21)', 0xff21, 'U'],
    ['ideographic space (3000)', 0x3000, 'U'],
    ['CJK Ext-B ideograph 𠀋 (surrogate pair)', 0x2000b, 'U'],
    ['vertical form comma FE10 (already vertical)', 0xfe10, 'U'],
    // R — rotated: Latin, digits, most Latin punctuation (the file default).
    ['A Latin capital', cp('A'), 'R'],
    ['z Latin small', cp('z'), 'R'],
    ['5 digit', cp('5'), 'R'],
    ['( ASCII paren', cp('('), 'R'],
    ['halfwidth full stop FF61', 0xff61, 'R'],
    // Tr — transform, fallback rotate: brackets, long vowel mark, quotes.
    ['ー prolonged sound mark (30FC)', 0x30fc, 'Tr'],
    ['「 left corner bracket (300C)', 0x300c, 'Tr'],
    ['」 right corner bracket (300D)', 0x300d, 'Tr'],
    ['（ fullwidth left paren (FF08)', 0xff08, 'Tr'],
    ['） fullwidth right paren (FF09)', 0xff09, 'Tr'],
    ['〈 angle bracket (3008)', 0x3008, 'Tr'],
    ['“ left double quote (201C)', 0x201c, 'Tr'],
    // Tu — transform, fallback upright: 、。, fullwidth ！？, small kana.
    ['、 ideographic comma (3001)', 0x3001, 'Tu'],
    ['。 ideographic full stop (3002)', 0x3002, 'Tu'],
    ['， fullwidth comma (FF0C)', 0xff0c, 'Tu'],
    ['！ fullwidth exclamation (FF01)', 0xff01, 'Tu'],
    ['？ fullwidth question (FF1F)', 0xff1f, 'Tu'],
    ['ぁ small hiragana a (3041)', 0x3041, 'Tu'],
    ['ッ small katakana tu (30C3)', 0x30c3, 'Tu'],
  ];
  for (const [label, c, expected] of cases) {
    it(`${label} → ${expected}`, () => {
      expect(verticalOrientation(c)).toBe(expected);
    });
  }

  it('resolves range boundaries (low−1, low, high) around the U CJK-Radicals block', () => {
    // 2E80..2E99 ; U (CJK Radicals Supplement). 2E7F is not listed ⇒ R default.
    expect(verticalOrientation(0x2e7f)).toBe('R'); // just below block start (unlisted → R)
    expect(verticalOrientation(0x2e80)).toBe('U'); // block start
    expect(verticalOrientation(0x2e99)).toBe('U'); // block end
    // CJK Unified Ideographs 4E00..9FFF ; U
    expect(verticalOrientation(0x4e00)).toBe('U'); // block start
    expect(verticalOrientation(0x9fff)).toBe('U'); // block end
  });

  it('resolves the Tr long-vowel mark exactly (30FC is Tr, neighbours are U)', () => {
    expect(verticalOrientation(0x30fb)).toBe('U'); // ・ katakana middle dot
    expect(verticalOrientation(0x30fc)).toBe('Tr'); // ー prolonged sound mark
    expect(verticalOrientation(0x30fd)).toBe('U'); // ヽ iteration mark
  });

  it('defaults unlisted code points to R (@missing default)', () => {
    expect(verticalOrientation(0x0041)).toBe('R'); // Latin A
    expect(verticalOrientation(0x2026)).toBe('R'); // … horizontal ellipsis (vo=R)
  });
});

describe('verticalFormSubstitute (UAX #50 §5 vertical-form glyph substitution)', () => {
  it('maps the CORNER-HANGING Tu punctuation (comma/full stop) to its U+FE1x form', () => {
    expect(verticalFormSubstitute(0x3001)).toBe(0xfe11); // 、 → ︑
    expect(verticalFormSubstitute(0x3002)).toBe(0xfe12); // 。 → ︒
    expect(verticalFormSubstitute(0xff0c)).toBe(0xfe10); // ， → ︐
  });

  it('does NOT substitute ！／？ (they stand upright & centred, not corner-hanging)', () => {
    // ！ FF01 → FE15 and ？ FF1F → FE16 are deliberately absent. The FE15/FE16
    // vertical forms are corner-designed in many fonts and shift ！／？ off the
    // column centre; the original fullwidth mark drawn upright is already the
    // correct, centred vertical appearance (PDF-verified on sample-26, #771).
    expect(verticalFormSubstitute(0xff01)).toBeNull(); // ！
    expect(verticalFormSubstitute(0xff1f)).toBeNull(); // ？
  });

  it('returns null for non-Tu punctuation (Tr substitutes via the OTHER map, R stays sideways)', () => {
    // ：； and 〖〗 are vo=Tr, so their vertical forms (FE13/FE14/FE17/FE18) live in
    // verticalBracketFormSubstitute (the Tr map), NOT here — this Tu-only map stays
    // null for them so the renderer never draws a Tr form on the upright(Tu) path.
    expect(verticalFormSubstitute(0xff1a)).toBeNull(); // ：
    expect(verticalFormSubstitute(0xff1b)).toBeNull(); // ；
    expect(verticalFormSubstitute(0x3016)).toBeNull(); // 〖
    expect(verticalFormSubstitute(0x3017)).toBeNull(); // 〗
    // … is vo=R: rotated sideways, its dots stack vertically already.
    expect(verticalFormSubstitute(0x2026)).toBeNull(); // …
  });

  it('returns null for code points without a vertical presentation form', () => {
    // Small kana are Tu but have no U+FExx form (font handles them via `vert`).
    expect(verticalFormSubstitute(0x3041)).toBeNull(); // ぁ
    // Corner brackets are Tr; verticalFormSubstitute is the Tu-only map (the Tr
    // bracket forms live in verticalBracketFormSubstitute), so this stays null.
    expect(verticalFormSubstitute(0x300c)).toBeNull(); // 「
    // Ordinary upright ideographs.
    expect(verticalFormSubstitute(cp('漢'))).toBeNull();
    // Latin.
    expect(verticalFormSubstitute(cp('A'))).toBeNull();
  });
});

describe('verticalBracketFormSubstitute (UAX #50 Tr code points → U+FE1x/FE3x vertical forms)', () => {
  it('maps each fullwidth bracket/paren/brace to its U+FE3x vertical form', () => {
    expect(verticalBracketFormSubstitute(0xff08)).toBe(0xfe35); // （ → ︵
    expect(verticalBracketFormSubstitute(0xff09)).toBe(0xfe36); // ） → ︶
    expect(verticalBracketFormSubstitute(0xff5b)).toBe(0xfe37); // ｛ → ︷
    expect(verticalBracketFormSubstitute(0xff5d)).toBe(0xfe38); // ｝ → ︸
    expect(verticalBracketFormSubstitute(0x3014)).toBe(0xfe39); // 〔 → ︹
    expect(verticalBracketFormSubstitute(0x3015)).toBe(0xfe3a); // 〕 → ︺
    expect(verticalBracketFormSubstitute(0x3010)).toBe(0xfe3b); // 【 → ︻
    expect(verticalBracketFormSubstitute(0x3011)).toBe(0xfe3c); // 】 → ︼
    expect(verticalBracketFormSubstitute(0x300a)).toBe(0xfe3d); // 《 → ︽
    expect(verticalBracketFormSubstitute(0x300b)).toBe(0xfe3e); // 》 → ︾
    expect(verticalBracketFormSubstitute(0x3008)).toBe(0xfe3f); // 〈 → ︿
    expect(verticalBracketFormSubstitute(0x3009)).toBe(0xfe40); // 〉 → ﹀
    expect(verticalBracketFormSubstitute(0x300c)).toBe(0xfe41); // 「 → ﹁
    expect(verticalBracketFormSubstitute(0x300d)).toBe(0xfe42); // 」 → ﹂
    expect(verticalBracketFormSubstitute(0x300e)).toBe(0xfe43); // 『 → ﹃
    expect(verticalBracketFormSubstitute(0x300f)).toBe(0xfe44); // 』 → ﹄
  });

  it('maps the vo=Tr fullwidth colon/semicolon and white lenticular brackets to their U+FE1x form (issue #969)', () => {
    // Word/PowerPoint ground truth (Yu Mincho tbRl + eaVert, PDF-adjudicated):
    // these vo=Tr code points SUBSTITUTE their U+FE1x vertical presentation form
    // and draw UPRIGHT — NOT the rotate fallback. The fullwidth colon/semicolon are
    // vo=Tr punctuation (not brackets) but share the same substitute-first behaviour,
    // so they live in this Tr map. FE13/FE14 decompose (UnicodeData `<vertical>`) to
    // the ASCII 003A/003B; vertical Japanese carries the FULLWIDTH forms, so — like
    // FE10 ← FF0C — the map keys on FF1A/FF1B. FE17/FE18 are the exact `<vertical>`
    // inverses of the white lenticular brackets 3016/3017.
    expect(verticalBracketFormSubstitute(0xff1a)).toBe(0xfe13); // ： → ︓ vertical colon
    expect(verticalBracketFormSubstitute(0xff1b)).toBe(0xfe14); // ； → ︔ vertical semicolon
    expect(verticalBracketFormSubstitute(0x3016)).toBe(0xfe17); // 〖 → ︗ vertical left white lenticular
    expect(verticalBracketFormSubstitute(0x3017)).toBe(0xfe18); // 〗 → ︘ vertical right white lenticular
  });

  it('returns null for Tr code points with no U+FE30 vertical form (ー, quotes)', () => {
    expect(verticalBracketFormSubstitute(0x30fc)).toBeNull(); // ー prolonged sound mark
    expect(verticalBracketFormSubstitute(0x201c)).toBeNull(); // “ left double quote
    expect(verticalBracketFormSubstitute(0x201d)).toBeNull(); // ” right double quote
  });

  it('returns null for non-bracket code points (Tu punctuation, ideographs, Latin)', () => {
    expect(verticalBracketFormSubstitute(0x3001)).toBeNull(); // 、 (Tu)
    expect(verticalBracketFormSubstitute(cp('漢'))).toBeNull(); // ideograph (U)
    expect(verticalBracketFormSubstitute(cp('A'))).toBeNull(); // Latin (R)
    expect(verticalBracketFormSubstitute(cp('('))).toBeNull(); // ASCII paren (R)
  });
});
