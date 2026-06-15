import { describe, it, expect } from 'vitest';
import { lineBoxHeight } from './renderer.js';
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

// ECMA-376 §17.6.5 / §17.3.1.32 only define docGrid line height for natural <=
// pitch. Word's runtime behaviour for taller lines (verified via pdftotext -bbox
// of the sample-9 Word PDF: a 20pt CJK title on a 20pt pitch is 40px = 2 cells)
// rounds East Asian lines UP to whole grid cells, while Latin-only lines keep
// their natural height above a one-cell floor (demo/sample-1's 18pt heading on
// an 18pt pitch stays ~natural, not 2 cells). The `eastAsian` flag gates it.
describe('lineBoxHeight — docGrid line-cell rounding (East Asian vs Latin)', () => {
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  it('rounds an East Asian line taller than the pitch UP to whole cells', () => {
    // glyph natural 22px > 20px pitch → ceil(22/20) = 2 cells = 40px.
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20, false, 0, true)).toBe(40);
  });
  it('keeps a Latin line at natural height (one-cell floor), NOT rounded to 2 cells', () => {
    // Same 22px natural, Latin → max(22, 20) = 22px (Word does not cell-round it).
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20, false, 0, false)).toBe(22);
  });
  it('puts a short East Asian line in a single cell (natural <= pitch)', () => {
    // glyph natural 13px <= 20px pitch → ceil(13/20) = 1 cell = 20px.
    expect(lineBoxHeight(null, 10.4, 2.6, 1, grid20, false, 0, true)).toBe(20);
  });
  it('rounds an East Asian line over two pitches up to three cells', () => {
    // glyph natural 41px → ceil(41/20) = 3 cells = 60px.
    expect(lineBoxHeight(null, 32.8, 8.2, 1, grid20, false, 0, true)).toBe(60);
  });
  it('does not cell-round when no docGrid is active (East Asian, off grid)', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, undefined, false, 0, true)).toBe(22);
  });
  it('defaults to Latin (no cell rounding) when the flag is omitted', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20)).toBe(22);
  });
});
