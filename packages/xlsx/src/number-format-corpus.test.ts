import { describe, it, expect } from 'vitest';
import { formatCellValue, formatCellValueWithColor } from './number-format.js';
import type { Cell, Styles } from './types.js';

// ────────────────────────────────────────────────────────────────────────────
// XL3 — corpus / table-driven fidelity suite for the §18.8.30 number-format
// grammar. Each row is `[value, formatCode, expected, source]`. `source` cites
// where the expected string comes from so no case rests on a guess:
//   S   = ECMA-376 Part 1 §18.8.30 worked example (page cited inline per block)
//   MS  = Microsoft "Number format codes" / TEXT-function documented example
//         (support.microsoft.com/office). These are the canonical Excel display
//         strings for the code + value.
//   XL  = de-facto Excel behaviour that both S and MS leave implicit but which
//         follows unambiguously from the placeholder rules already in §18.8.30
//         (e.g. `#` dropping insignificant zeros, `0` padding them). Only used
//         where S/MS give the *rule* but not that exact value.
//   CORPUS = a real formatCode harvested from a private sample workbook's
//         styles.xml (see the XL3 task brief); the expected string is derived
//         from the same §18.8.30 rules, cross-checked against how Excel renders
//         that cell.
// ────────────────────────────────────────────────────────────────────────────

const FMT_ID = 164;

function styles(formatCode: string): Styles {
  return {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: FMT_ID, alignH: null, alignV: null, wrapText: false }],
    numFmts: [{ numFmtId: FMT_ID, formatCode }],
    dxfs: [],
  };
}
const numCell = (n: number): Cell => ({ row: 1, col: 1, value: { type: 'number', number: n }, styleIndex: 0 });
const fmt = (n: number, code: string) => formatCellValue(numCell(n), styles(code));

type Row = [value: number, code: string, expected: string, source: string];

function table(name: string, rows: Row[]) {
  describe(name, () => {
    for (const [value, code, expected, source] of rows) {
      it(`${JSON.stringify(code)} @ ${value} → ${JSON.stringify(expected)} [${source}]`, () => {
        expect(fmt(value, code)).toBe(expected);
      });
    }
  });
}

// ── Digit placeholders: 0 / # / ? and the decimal point ─────────────────────
// §18.8.30 "Decimal places and significant digits" table (p.1786):
//   ####.#   1234.59 → 1234.6
//   #.000    8.9     → 8.900
//   0.#      .631    → 0.6
//   #.0#     12      → 12.0     |  1234.568 → 1234.57
table('placeholders 0 / # / . (§18.8.30 p.1786 table)', [
  [1234.59, '####.#', '1234.6', 'S'],
  [8.9, '#.000', '8.900', 'S'],
  [0.631, '0.#', '0.6', 'S'],
  [12, '#.0#', '12.0', 'S'],
  [1234.568, '#.0#', '1234.57', 'S'],
  // "If the format contains only number signs (#) to the left of the decimal
  // point, numbers less than 1 begin with a decimal point." (§18.8.30 p.1786)
  [0.5, '#.##', '.5', 'S'],
  // 0 to the left pads the integer part with a leading zero. (rule, §18.8.30)
  [0.5, '0.##', '0.5', 'XL'],
  // # drops insignificant zeros; 8.9 under #.## shows 8.9. (§18.8.30 p.1776)
  [8.9, '#.##', '8.9', 'S'],
  // 0 shows insignificant zeros: 8.9 under #.00 → 8.90. (§18.8.30 p.1776)
  [8.9, '#.00', '8.90', 'S'],
  // extra integer digits are always shown even with a single placeholder.
  [1234.568, '0', '1235', 'XL'],
]);

// ── Thousands separator and comma scaling ───────────────────────────────────
// §18.8.30 "Display a thousands separator" table (p.1787):
//   #,###   12000    → 12,000
//   #,      12000    → 12
//   0.0,,   12200000 → 12.2
// and "Display both ..." examples.
table('thousands separator & comma scaling (§18.8.30 p.1787)', [
  [12000, '#,###', '12,000', 'S'],
  [12000, '#,', '12', 'S'],
  [12200000, '0.0,,', '12.2', 'S'],
  [1234567, '#,##0', '1,234,567', 'MS'],
  [1234.5, '#,##0.0', '1,234.5', 'MS'],
  // "A comma that follows a placeholder scales the number by one thousand."
  // #.0,, → 12,200,000 shows 12.2 (§18.8.30 p.1784 comma row).
  [12200000, '#.0,,', '12.2', 'S'],
  // A single trailing comma scales by one thousand: 12,200,000 → 12,200.
  [12200000, '#,##0,', '12,200', 'XL'],
]);

// ── Percent (§18.8.30) ──────────────────────────────────────────────────────
// "multiply the number by 100 and add the percentage symbol" — .08 → 8%,
// 2.8 → 280% (§18.8.30 p.1784 & p.1788).
table('percent (§18.8.30 p.1788)', [
  [0.08, '0%', '8%', 'S'],
  [2.8, '0%', '280%', 'S'],
  [0.5, '0%', '50%', 'MS'],
  [0.1234, '0.0%', '12.3%', 'MS'],
]);

// ── Scientific notation (§18.8.30) ──────────────────────────────────────────
// "If the format is 0.00E+00, and the value 12,200,000 is in the cell, the
// number 1.22E+07 is displayed. If the number format is #0.0E+0, then the
// number 12.2E+6 is displayed." (§18.8.30 p.1785)
table('scientific (§18.8.30 p.1785)', [
  [12200000, '0.00E+00', '1.22E+07', 'S'],
  [12200000, '#0.0E+0', '12.2E+6', 'S'],
  // E- suppresses the + on positive exponents (§18.8.30 p.1788).
  [12200000, '0.00E-00', '1.22E07', 'XL'],
]);

// ── Fractions (§18.8.30) ────────────────────────────────────────────────────
// §18.8.30 p.1776 built-in table: id 12 = "# ?/?", id 13 = "# ??/??".
// §18.8.30 p.1787 shows 5.25 → "5 1/4" and 5.3 → "5 3/10" with "# ???/???"
// (that table collapses the alignment spaces for readability). The `?`
// placeholder is defined to "add spaces for insignificant zeros ... so that
// [values] align" (§18.8.30 p.1786), so under `???` Excel right-aligns the
// numerator (pad left) and left-aligns the denominator (pad right) inside the
// 3-wide fields — the exact space counts follow mechanically from that rule.
table('fractions (§18.8.30 built-in ids 12/13 & p.1786-1787 ? rule)', [
  // 5.25: numerator "1" in a 3-wide `?` field = "  1"; denominator "4" = "4  ".
  [5.25, '# ???/???', '5   1/4  ', 'XL'],
  [5.3, '# ???/???', '5   3/10 ', 'XL'],
  // # ?/? — single-digit fields. 0.5 → " 1/2" (numerator padded to 1). (MS)
  [0.5, '# ?/?', ' 1/2', 'MS'],
  // MS "Custom number format" fraction example: 0.25 as "# ?/?" → " 1/4".
  [0.25, '# ?/?', ' 1/4', 'MS'],
  // Fixed denominator: ?/8 renders in eighths. 0.25 → "2/8". (MS docs)
  [0.25, '?/8', '2/8', 'MS'],
  // Whole number under a mixed-fraction format blanks the " n/d" group so the
  // integer aligns with fractional neighbours (§18.8.30 `?` alignment). For
  // "# ?/?" that is one space + "?" + "/" + "?" = 4 blank columns after "4".
  [4, '# ?/?', '4    ', 'XL'],
]);

// ── Sign sections (§18.8.31 "Number Format Codes") ──────────────────────────
// "define the formats for positive numbers, negative numbers, zero values, and
// text, in that order." Two sections: first = positive & zero, second =
// negative. One section = all numbers. (§18.8.31 p.1783)
table('sign sections (§18.8.31 p.1783)', [
  [5, '0;(0);"-"', '5', 'S'],
  [-5, '0;(0);"-"', '(5)', 'S'],
  [0, '0;(0);"-"', '-', 'S'],
  // one section applies to all numbers (negative keeps its sign).
  [-5, '0.0', '-5.0', 'XL'],
  // two sections: first for positive & zero, second for negative.
  [0, '0.0;(0.0)', '0.0', 'S'],
  [-3, '0.0;(0.0)', '(3.0)', 'S'],
  // §18.8.30 p.1785 "Display both text and numbers" worked example:
  //   $0.00" Surplus";$-0.00" Shortage"  → 125.74 "$125.74 Surplus"
  //   and -125.74 "$-125.74 Shortage".  The negative section keeps its own
  //   literal minus, so the magnitude carries no extra sign.
  [125.74, '$0.00" Surplus";$-0.00" Shortage"', '$125.74 Surplus', 'S'],
  [-125.74, '$0.00" Surplus";$-0.00" Shortage"', '$-125.74 Shortage', 'S'],
  // CORPUS sample-4 numeric flag: positive → literal text, others blank.
  // The positive section is a bare quoted literal (no digit placeholder), so a
  // positive value shows only the literal; zero/negative sections are empty.
  [1, '"再発注";"";""', '再発注', 'CORPUS'],
  [0, '"再発注";"";""', '', 'CORPUS'],
]);

// ── Conditional sections (§18.8.30 "Specify conditions") ────────────────────
// "enclose the condition in square brackets ... comparison operator and a
// value." The example [Red][<=100];[Blue][>100] colours by threshold. Here we
// test the *section selection* driven by the condition (colour tested below).
table('conditional sections (§18.8.30 p.1787)', [
  // [>=1000]#,##0,"K";0  — thousands get a K suffix, else plain.
  [2500, '[>=1000]#,##0,"K";0', '3K', 'XL'],
  [999, '[>=1000]#,##0,"K";0', '999', 'XL'],
  // CORPUS sample-8 phone masks (real styles.xml):
  //   [<=999]000;[<=9999]000\-00;000\-0000
  // Each `0` forces a digit, so shorter numbers zero-pad to fill every
  // placeholder position and the literal `-` stays fixed. 1234 fills the 5
  // placeholders of `000-00` as "01234" → "012-34"; 123456 fills the 7
  // placeholders of `000-0000` as "0123456" → "012-3456".
  [12, '[<=999]000;[<=9999]000\\-00;000\\-0000', '012', 'CORPUS'],
  [1234, '[<=999]000;[<=9999]000\\-00;000\\-0000', '012-34', 'CORPUS'],
  [123456, '[<=999]000;[<=9999]000\\-00;000\\-0000', '012-3456', 'CORPUS'],
  // ── Sign handling in condition-selected sections ──────────────────────────
  // A section selected by its *matching condition* formats the value's
  // magnitude; the section's own literals carry the sign presentation. This is
  // the same model §18.8.30 uses for the positional negative section: the
  // spec's own example `$0.00" Surplus";$-0.00" Shortage"` (p.1785) renders
  // -125.74 as "$-125.74 Shortage" — magnitude + the section's literal `-` —
  // and built-ins 37-40 (`#,##0 ;(#,##0)`) put parentheses where the minus
  // would be. Applying the sign on top would double it ("--5.0",
  // "-(1,234.50)"), which Excel does not do (independently verified against
  // Excel in the PR #799 review).
  [-5, '[>0]0.0;[<0]\\-0.0', '-5.0', 'XL'],
  [-1234.5, '[>=0]#,##0.00;[<0](#,##0.00)', '(1,234.50)', 'XL'],
  // No sign literal in the condition-matched section → the sign is simply not
  // rendered (the section owns the sign presentation and provides none) —
  // mechanically the same rule that makes positional `0;(0)` show -5 as "(5)"
  // rather than re-adding a minus.
  [-5, '[<0]0.0;0.0', '5.0', 'XL'],
  // The unconditional "else" section is the fallback for values no condition
  // claimed. It mirrors the positional fallback rule (§18.8.31: a negative
  // formatted by the only/positive section keeps its sign, cf. `0.0` @ -5 →
  // "-5.0" above), so the sign survives here.
  [-500, '[>=1000]#,##0,"K";0', '-500', 'XL'],
]);

// ── Currency / locale-currency `[$sym-LCID]` (§18.8.30 p.1791) ──────────────
// "[$<Currency String>-<language info>]" — the currency string between $ and -
// is emitted literally; the -LCID tail is metadata and dropped.
table('locale currency [$sym-LCID] (§18.8.30 p.1791)', [
  // CORPUS sample-25: [$¥-411]#,##0.00 → ¥ prefix, LCID dropped.
  [1234.5, '[$¥-411]#,##0.00', '¥1,234.50', 'CORPUS'],
  // $ in a [$...] wrapper is a literal currency glyph, not a section symbol.
  [1234.5, '[$USD-409]#,##0.00', 'USD1,234.50', 'S'],
  // €/EUR locale currency (task brief example [$€-407]).
  [99, '[$€-407]#,##0.00', '€99.00', 'CORPUS'],
  // System long-date currency-syntax metadata [$-F800] is a locale tag only.
  // (date rendering of [$-F800] handled by the date path.)
]);

// ── Literals, escapes, quotes, spacing (§18.8.30 p.1784-1786) ───────────────
table('literals / escapes / spacing (§18.8.30 p.1784-1786)', [
  // \ displays the next char: 0\! at 3 → "3!" (§18.8.30 p.1785)
  [3, '0\\!', '3!', 'S'],
  // (000) at 12 → "(012)" — parens & leading zeros literal. (§18.8.30 p.1784)
  [12, '(000)', '(012)', 'S'],
  // "text" literal: 0.00 "dollars" at 1.23 → "1.23 dollars" (§18.8.30 p.1776)
  [1.23, '0.00 "dollars"', '1.23 dollars', 'S'],
  // quoted literal around the number.
  [3, '0" units"', '3 units', 'XL'],
  // _ skips the width of the next char (renders a space in a proportional grid).
  // _(0.0_) at 2.3 → " 2.3 " (leading + trailing space of a "(" / ")"). (p.1786)
  [2.3, '_(0.0_)', ' 2.3 ', 'S'],
]);

// ── Real accounting / currency formats harvested from private sample
//    workbooks' styles.xml (XL3 task brief). Expected strings are the display
//    Excel shows for that cell — `_(` pads a `(`-width space, `* ` fills, the
//    `"-"??` zero section shows a dash with two alignment spaces. ────────────
table('accounting & currency (CORPUS — private sample styles.xml)', [
  // Standard "Accounting" format (Excel built-in, ubiquitous in the samples).
  [1234.5, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', '  1,234.50 ', 'CORPUS'],
  [-1234.5, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', '  (1,234.50)', 'CORPUS'],
  [0, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', '  -   ', 'CORPUS'],
  // Yen currency with an explicit negative section (sample-1/4/7/25).
  [1234.5, '"¥"#,##0.00;"¥"\\-#,##0.00', '¥1,234.50', 'CORPUS'],
  [-1234.5, '"¥"#,##0.00;"¥"\\-#,##0.00', '¥-1,234.50', 'CORPUS'],
  // Dollar currency, single section (sample-1/7/25).
  [1234.5, '"$"#,##0.00', '$1,234.50', 'CORPUS'],
  // Parenthesised-negative integer with `_)` alignment (sample-2/6).
  [1234, '#,##0_);\\(#,##0\\)', '1,234 ', 'CORPUS'],
  [-1234, '#,##0_);\\(#,##0\\)', '(1,234)', 'CORPUS'],
  // Locale-currency yen with negative sign outside the bracket (sample-25).
  [123, '[$¥-411]#,##0.00;\\-[$¥-411]#,##0.00', '¥123.00', 'CORPUS'],
  [-123, '[$¥-411]#,##0.00;\\-[$¥-411]#,##0.00', '-¥123.00', 'CORPUS'],
  // Trailing `_ ` alignment space after a plain 2-decimal number (sample-26).
  [1234.5, '#,##0.00_ ', '1,234.50 ', 'CORPUS'],
]);

// ── More placeholder / decimal edge cases grounded in the §18.8.30 rules ────
table('decimal & rounding edge cases (§18.8.30 rules)', [
  // 0.000 shows three forced decimals (sample-30 real format).
  [5, '0.000', '5.000', 'CORPUS'],
  // 0.0 rounds to one decimal (sample-30).
  [3.14159, '0.0', '3.1', 'CORPUS'],
  // #.## drops trailing insignificant zeros but keeps significant ones.
  [1234.5, '#.##', '1234.5', 'S'],
  // #,##0 rounds to an integer with grouping.
  [1234.567, '#,##0', '1,235', 'MS'],
  // Rounding up carries into the integer part.
  [9.99, '0.0', '10.0', 'XL'],
  // Negative single-section keeps the minus in front of the grouped number.
  [-1234.5, '#,##0.00', '-1,234.50', 'XL'],
]);

// ── Text section (§18.8.30 "Include a section for text entry", p.1786) ──────
describe('text section (§18.8.30 p.1786)', () => {
  const textCell = (t: string): Cell => ({ row: 1, col: 1, value: { type: 'text', text: t }, styleIndex: 0 });
  const ft = (t: string, code: string) => formatCellValue(textCell(t), styles(code));
  it('@ substitutes the typed text ["gross receipts for "@ ]', () => {
    // §18.8.30 p.1786 worked example.
    expect(ft('June', '"gross receipts for "@')).toBe('gross receipts for June');
  });
  it('"Bob "@" Smith" wraps the text [§18.8.30 p.1786 @ row]', () => {
    expect(ft('John', '"Bob "@" Smith"')).toBe('Bob John Smith');
  });
  it('empty text section hides the value [;;; hides all — §18.8.31]', () => {
    expect(ft('anything', ';;;')).toBe('');
  });
  it('a 3-section format has no text section, so text is unaffected [§18.8.30 p.1786]', () => {
    // "再発注";"";"" is CORPUS sample-4's numeric flag (positive → 再発注,
    // negative/zero → blank). With only three sections there is no text
    // section, so a text value passes through untouched.
    expect(ft('x', '"再発注";"";""')).toBe('x');
  });
  it('a format without a text section leaves text unaffected [§18.8.30 p.1786]', () => {
    expect(ft('hello', '0.00')).toBe('hello');
  });
});

// ── Colors — [Red] etc. threaded through formatCellValueWithColor ───────────
// §18.8.30 "Specify colors" (p.1787): eight named colours + [ColorN] index.
describe('section colors (§18.8.30 p.1787)', () => {
  const wc = (n: number, code: string) => formatCellValueWithColor(numCell(n), styles(code));
  it('[Red] on the negative section colours only negatives', () => {
    // CORPUS sample-11/25/30: "¥"#,##0;[Red]"¥"\-#,##0
    expect(wc(-5, '"¥"#,##0;[Red]"¥"\\-#,##0')).toEqual({ text: '¥-5', color: '#FF0000' });
    expect(wc(5, '"¥"#,##0;[Red]"¥"\\-#,##0')).toEqual({ text: '¥5' });
  });
  it('named colours map to their palette hex [§18.8.30 p.1787 list]', () => {
    expect(wc(1, '[Blue]0').color).toBe('#0000FF');
    expect(wc(1, '[Green]0').color).toBe('#008000');
    expect(wc(1, '[Magenta]0').color).toBe('#FF00FF');
  });
  it('[ColorN] uses the legacy indexed palette: [Color3] = Red [§18.8.30 note p.1788]', () => {
    // "[Color1] refers to indexed=8 (black) ... [Color3] for Red."
    expect(wc(1, '[Color3]0').color).toBe('#FF0000');
    expect(wc(1, '[Color1]0').color).toBe('#000000');
  });
  it('condition + colour: [Red][<=100];[Blue][>100] [§18.8.30 p.1787 example]', () => {
    expect(wc(50, '[Red][<=100];[Blue][>100]').color).toBe('#FF0000');
    expect(wc(150, '[Red][<=100];[Blue][>100]').color).toBe('#0000FF');
  });
  it('no colour token → color is undefined', () => {
    expect(wc(1, '0.00').color).toBeUndefined();
  });
});
