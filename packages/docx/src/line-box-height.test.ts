import { describe, it, expect } from 'vitest';
import { lineBoxHeight } from './renderer.js';
import type { LineSpacing } from './types.js';

// Regression: some generators emit <w:spacing w:line="0" w:lineRule="exact"/> on
// table cells (e.g. private/sample-7.docx). A literal exact-0 line collapsed every
// line to 0px, so table rows fell back to the 10px minimum and cell content
// overlapped. Word ignores a zero line value and uses single spacing; lineBoxHeight
// must do the same (ECMA-376 §17.3.1.33).

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
    expect(lineBoxHeight(exact(240), 12, 3, 1, undefined)).toBe(240);
  });
  it('still applies the auto multiplier for a positive value', () => {
    expect(lineBoxHeight(auto(2), 12, 3, 1, undefined)).toBe(30); // natural 15 × 2
  });
  it('unspecified spacing is single (natural)', () => {
    expect(lineBoxHeight(null, 12, 3, 1, undefined)).toBe(15);
  });
});
