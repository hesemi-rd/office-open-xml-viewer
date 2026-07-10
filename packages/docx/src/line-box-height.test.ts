import { describe, it, expect } from 'vitest';
import { docGridLineCells, isGridLineRule, lineBoxHeight } from './line-layout.js';
import type { LineSpacing } from './types.js';

// Regression: some generators emit <w:spacing w:line="0" w:lineRule="exact"/> on
// table cells (e.g. private/sample-7.docx). A literal exact-0 line collapsed every
// line to 0px, so table rows fell back to the 10px minimum and cell content
// overlapped. ECMA-376 §17.3.1.33 leaves a zero line undefined; Word's native
// model ([MS-DOC] LSPD) cannot represent an exact 0 (exact = negative dyaLine)
// and resolves a non-negative dyaLine as max(dyaLine, single spacing) — so 0
// means exactly single spacing. lineBoxHeight must do the same.

const exact = (value: number): LineSpacing => ({ value, rule: 'exact', explicit: true });
const auto = (value: number): LineSpacing => ({ value, rule: 'auto', explicit: true });
const atLeast = (value: number): LineSpacing => ({
  value,
  rule: 'atLeast',
  explicit: true,
});

describe('lineBoxHeight — degenerate zero line spacing', () => {
  // ascent 12 + descent 3 = 15px natural single-line height (no grid).
  it('treats exact line=0 as single spacing (natural), not 0', () => {
    expect(lineBoxHeight(exact(0), 12, 3, 1, undefined)).toBe(15);
  });
  it('treats auto value=0 as single spacing (natural), not 0', () => {
    expect(lineBoxHeight(auto(0), 12, 3, 1, undefined)).toBe(15);
  });
  it('still honors a positive exact line value', () => {
    // 12pt exact (w:line="240" twips / 20 = 12pt; LineSpacing.value is pt).
    expect(lineBoxHeight(exact(12), 12, 3, 1, undefined)).toBe(12);
  });
  it('treats negative exact/auto values as single spacing too', () => {
    expect(lineBoxHeight(exact(-5), 12, 3, 1, undefined)).toBe(15);
    expect(lineBoxHeight(auto(-1), 12, 3, 1, undefined)).toBe(15);
  });
  it('snaps a degenerate line to the grid pitch in docGrid sections', () => {
    // Same fallback as unspecified spacing: on-grid, the pitch governs.
    expect(
      lineBoxHeight(exact(0), 12, 3, 1, { type: 'lines', linePitchPt: 18 }),
    ).toBe(18);
  });
  it('still applies the auto multiplier for a positive value', () => {
    expect(lineBoxHeight(auto(2), 12, 3, 1, undefined)).toBe(30); // natural 15 × 2
  });
  it('unspecified spacing is single (natural)', () => {
    expect(lineBoxHeight(null, 12, 3, 1, undefined)).toBe(15);
  });
});

describe('lineBoxHeight — snapToChars line-grid participation', () => {
  const grid20 = { type: 'snapToChars', linePitchPt: 20 } as const;

  it('recognizes snapToChars as an active line-grid rule', () => {
    expect(isGridLineRule(grid20)).toBe(true);
  });

  it('applies the line pitch as well as the character-grid behavior', () => {
    expect(lineBoxHeight(null, 8, 2, 1, grid20, false, 0, true)).toBe(20);
  });
});

describe('lineBoxHeight — atLeast with an active line grid', () => {
  const grid18 = { type: 'lines', linePitchPt: 18 } as const;
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  it('retains the active grid minimum when the authored minimum is smaller', () => {
    expect(lineBoxHeight(atLeast(18), 10, 2, 1, grid20)).toBe(20);
  });

  it('retains the authored minimum when it is larger than the grid minimum', () => {
    expect(lineBoxHeight(atLeast(24), 10, 2, 1, grid20)).toBe(24);
  });

  // BEHAVIOUR PIN, not Word-verified ground truth. For an East Asian atLeast
  // line the grid-minimum term is the em-based single cell (docGridLineCells),
  // no longer the glyph-box cell rounding: a 12pt line whose substituted glyph
  // box is 19.22px on an 18pt pitch resolves to its natural 19.22px (max of
  // natural / authored 18 / em cell 18) — previously 36px (glyph box rounded to
  // 2 cells). No corpus sample exercises atLeast inside an active line grid and
  // a Word fixture export could not be captured when this was pinned, so Word's
  // exact atLeast-on-grid height is UNVERIFIED; the pin makes any future change
  // to this resolution deliberate rather than accidental. Ruby lines keep the
  // measured-glyph-box minimum (they reserve real furigana height).
  it('[pin] EA atLeast takes max(natural, authored, em single cell) — unverified vs Word', () => {
    expect(lineBoxHeight(atLeast(18), 15.38, 3.84, 1, grid18, false, 0, true, 12)).toBeCloseTo(19.22, 6);
  });
  it('[pin] EA atLeast still snaps to the em cell when it exceeds natural and authored', () => {
    // em 20 on pitch 18 → 2 cells = 36 > natural 12 and authored 18.
    expect(lineBoxHeight(atLeast(18), 10, 2, 1, grid18, false, 0, true, 20)).toBe(36);
  });
  it('[pin] a RUBY EA atLeast line keeps the measured glyph-box minimum', () => {
    // glyph box 41px (base + furigana reserve) → ceil(41/18) = 3 cells = 54.
    expect(lineBoxHeight(atLeast(18), 33, 8, 1, grid18, true, 0, true, 13.5)).toBe(54);
  });
});

// ECMA-376 §17.6.5 / §17.3.1.32 define the line pitch as the height of ONE
// single-spaced line; how many whole grid CELLS a single-spaced East Asian line
// occupies is Word runtime behaviour. That cell count is a function of the run's
// EM (font size) — NOT the substituted-font glyph box (fontBoundingBox, ~1.6em),
// which over-counts every EA line whose em is comfortably inside one pitch.
//
// Word-PDF ground truth (pdftotext -bbox):
//   • sample-35: docGrid pitch 18pt; a 12pt centred CJK heading is 1 cell (18pt),
//     and the 10.5pt body lines are 1 cell — even though 12×1.6 = 19.2 > 18.
//   • sample-9 : docGrid pitch 20pt; a 20pt CJK title is 2 cells (40pt).
// `ceil(em/pitch)` fits the first two but fails the 20-on-20 boundary (→1, want
// 2). `floor(em/pitch)+1` fits ALL three: a single-spaced EA line whose em-square
// exactly fills k pitches still needs its inter-line leading, spilling into the
// (k+1)-th cell. Latin-only lines are NOT cell-rounded (§17.6.5 leaves them at
// default spacing); the `eastAsian` flag gates the rule.
describe('docGridLineCells — em-based East Asian grid cell count (§17.6.5)', () => {
  it('a sub-pitch em occupies a single cell (12pt heading / 18pt pitch → 1)', () => {
    expect(docGridLineCells(12, 18)).toBe(1);
  });
  it('a 13pt em on an 18pt pitch is still a single cell', () => {
    expect(docGridLineCells(13, 18)).toBe(1);
  });
  it('a 10.5pt body em on an 18pt pitch is a single cell', () => {
    expect(docGridLineCells(10.5, 18)).toBe(1);
  });
  it('an em equal to the pitch occupies TWO cells (20pt / 20pt → 2)', () => {
    expect(docGridLineCells(20, 20)).toBe(2);
  });
  it('the pitch boundary rounds up (18pt em / 18pt pitch → 2)', () => {
    expect(docGridLineCells(18, 18)).toBe(2);
  });
  it('an em between one and two pitches occupies two cells', () => {
    expect(docGridLineCells(30, 20)).toBe(2);
  });
  it('an em at exactly two pitches occupies three cells', () => {
    expect(docGridLineCells(40, 20)).toBe(3);
  });
});

describe('lineBoxHeight — docGrid line-cell rounding (East Asian vs Latin)', () => {
  const grid18 = { type: 'lines', linePitchPt: 18 } as const;
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  // Yu Mincho reports fontBoundingBox ≈ 1.602 em, so a 12pt run's glyph box is
  // ~19.22px (asc 15.38 + desc 3.84) — over the 18pt pitch. The OLD rule
  // ceil(19.22/18) = 2 cells over-counted; the em (12pt) → floor(12/18)+1 = 1
  // cell = 18px is what Word renders (sample-35 heading).
  it('snaps a sub-pitch EA line to ONE cell by its em, not its inflated glyph box', () => {
    expect(lineBoxHeight(null, 15.38, 3.84, 1, grid18, false, 0, true, 12)).toBe(18);
  });
  it('rounds an EA line whose em equals the pitch UP to two cells (20/20 → 40)', () => {
    // glyph box 32.04px (25.63 + 6.41); em 20 on pitch 20 → floor(20/20)+1 = 2.
    expect(lineBoxHeight(null, 25.63, 6.41, 1, grid20, false, 0, true, 20)).toBe(40);
  });
  it('rounds an EA line whose em exceeds the pitch to two cells (30/20 → 40)', () => {
    expect(lineBoxHeight(null, 24, 6, 1, grid20, false, 0, true, 30)).toBe(40);
  });
  it('keeps a Latin line at natural height (one-cell floor), NOT cell-rounded', () => {
    // 22px natural, Latin → max(22, 20) = 22px (Word does not cell-round it).
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20, false, 0, false, 22)).toBe(22);
  });
  it('a RUBY EA line reserves its measured furigana height, NOT the em cell count', () => {
    // sample-5: a 13.5pt ruby base + 8pt rt on an 18pt pitch reserves the base +
    // annotation glyph box (~41px = asc 33 + desc 8), which Word spreads over
    // whole cells (ceil(41/18) = 3 cells = 54px). The em (13.5pt) would collapse
    // it to one cell (18px) and clip the furigana, so the em rule must NOT apply
    // to ruby lines — hasRuby (arg 6) keeps them on the measured glyph box.
    expect(lineBoxHeight(null, 33, 8, 1, grid18, true, 0, true, 13.5)).toBe(54);
  });
  it('does not cell-round when no docGrid is active (East Asian, off grid)', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, undefined, false, 0, true, 22)).toBe(22);
  });
  it('defaults to Latin (no cell rounding) when the flag is omitted', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20)).toBe(22);
  });
});
