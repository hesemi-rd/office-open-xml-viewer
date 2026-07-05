import { describe, it, expect } from 'vitest';
import { parseFieldFormatSwitch } from './field-format-switch';

describe('parseFieldFormatSwitch — ECMA-376 §17.16.4.3.1 general-formatting switch', () => {
  it('returns null when the instruction carries no \\* switch', () => {
    expect(parseFieldFormatSwitch('PAGE')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* MERGEFORMAT')).toBeNull(); // MERGEFORMAT is not a number format
    expect(parseFieldFormatSwitch('')).toBeNull();
  });

  it('maps the numeric switch arguments to ST_NumberFormat values (case-sensitive)', () => {
    // Arabic -> decimal; ArabicDash -> "- N -"; Hex -> uppercase hexadecimal
    expect(parseFieldFormatSwitch('PAGE \\* Arabic')).toBe('decimal');
    expect(parseFieldFormatSwitch('PAGE \\* ArabicDash')).toBe('numberInDash');
    expect(parseFieldFormatSwitch('PAGE \\* Hex')).toBe('hex');
    // Roman (uppercase) vs roman (lowercase) are DISTINCT arguments
    expect(parseFieldFormatSwitch('PAGE \\* Roman')).toBe('upperRoman');
    expect(parseFieldFormatSwitch('PAGE \\* roman')).toBe('lowerRoman');
    // ALPHABETIC vs alphabetic
    expect(parseFieldFormatSwitch('PAGE \\* ALPHABETIC')).toBe('upperLetter');
    expect(parseFieldFormatSwitch('PAGE \\* alphabetic')).toBe('lowerLetter');
  });

  it('ignores a trailing MERGEFORMAT after a real format switch', () => {
    expect(parseFieldFormatSwitch('PAGE \\* roman \\* MERGEFORMAT')).toBe('lowerRoman');
    expect(parseFieldFormatSwitch('PAGE \\* MERGEFORMAT \\* Roman')).toBe('upperRoman');
  });

  it('tolerates extra whitespace and surrounding switches', () => {
    expect(parseFieldFormatSwitch('  PAGE    \\*    Roman   ')).toBe('upperRoman');
    expect(parseFieldFormatSwitch('PAGE \\* roman \\# 0')).toBe('lowerRoman');
  });

  it('maps the locale-independent international switch arguments', () => {
    // §17.16.4.3.1 — each of these rows maps to exactly ONE ST_NumberFormat
    // regardless of the field language, so they are safe to route.
    expect(parseFieldFormatSwitch('PAGE \\* ARABICABJAD')).toBe('arabicAbjad');
    expect(parseFieldFormatSwitch('PAGE \\* ARABICALPHA')).toBe('arabicAlpha');
    expect(parseFieldFormatSwitch('PAGE \\* HEBREW1')).toBe('hebrew1');
    expect(parseFieldFormatSwitch('PAGE \\* HEBREW2')).toBe('hebrew2');
    expect(parseFieldFormatSwitch('PAGE \\* HINDIARABIC')).toBe('hindiNumbers');
    expect(parseFieldFormatSwitch('PAGE \\* HINDILETTER1')).toBe('hindiVowels');
    expect(parseFieldFormatSwitch('PAGE \\* HINDILETTER2')).toBe('hindiConsonants');
    expect(parseFieldFormatSwitch('PAGE \\* THAIARABIC')).toBe('thaiNumbers');
    expect(parseFieldFormatSwitch('PAGE \\* THAILETTER')).toBe('thaiLetters');
    expect(parseFieldFormatSwitch('PAGE \\* CHOSUNG')).toBe('chosung');
    expect(parseFieldFormatSwitch('PAGE \\* GANADA')).toBe('ganada');
    expect(parseFieldFormatSwitch('PAGE \\* DBCHAR')).toBe('decimalFullWidth');
    expect(parseFieldFormatSwitch('PAGE \\* SBCHAR')).toBe('decimalHalfWidth');
  });

  it('returns null for a format argument this converter does not support', () => {
    // CardText etc. are recognised switches but not numeric-native — the caller
    // then keeps the section fmt / decimal. We surface null (not decimal) so the
    // caller can distinguish "no override" from "override to decimal".
    expect(parseFieldFormatSwitch('PAGE \\* CardText')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* Ordinal')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* BAHTTEXT')).toBeNull(); // Thai spell-out
  });

  it('returns null for the LOCALE-DEPENDENT switch arguments (needs w:lang)', () => {
    // §17.16.4.3.1 — these map to different ST_NumberFormat values per language
    // (e.g. DBNUM1 → ideographDigital in ja-JP vs koreanDigital in ko-KR). We do
    // not thread the field language, so we decline rather than guess a script.
    expect(parseFieldFormatSwitch('PAGE \\* CHINESENUM1')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* CHINESENUM2')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* CHINESENUM3')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* DBNUM1')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* DBNUM2')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* KANJINUM1')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* IROHA')).toBeNull();
    expect(parseFieldFormatSwitch('PAGE \\* AIUEO')).toBeNull();
  });
});
