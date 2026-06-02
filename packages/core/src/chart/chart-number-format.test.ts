import { describe, it, expect } from 'vitest';
import { formatChartVal, formatChartValWithCode } from './chart-number-format.js';

describe('formatChartVal (General)', () => {
  it('shows integers raw with no abbreviation', () => {
    expect(formatChartVal(7507)).toBe('7507');
    expect(formatChartVal(0)).toBe('0');
  });
  it('trims trailing decimal zeros', () => {
    expect(formatChartVal(0.5)).toBe('0.5');
    expect(formatChartVal(1.25)).toBe('1.25');
  });
});

describe('formatChartValWithCode — thousands separator', () => {
  it('adds commas with #,##0', () => {
    expect(formatChartValWithCode(1234567, '#,##0')).toBe('1,234,567');
    expect(formatChartValWithCode(7507, '#,##0')).toBe('7,507');
  });
  it('keeps decimals alongside the separator', () => {
    expect(formatChartValWithCode(1234.5, '#,##0.0')).toBe('1,234.5');
  });
});

describe('formatChartValWithCode — sample-2 slide-16 data labels (§18.8.30)', () => {
  // The bar-chart series number format is `#,##0_);[Red]\(#,##0\)`. Before the
  // parser wired `valFormatCode` through, these labels rendered as "7507" with
  // no comma; PowerPoint shows "7,507". The positive section's `_)` reserves a
  // glyph-width of ')' so the value aligns with parenthesised negatives — we
  // render it as a trailing space, matching Excel.
  const CODE = '#,##0_);[Red]\\(#,##0\\)';
  it('positive value gets a comma (the reported bug)', () => {
    expect(formatChartValWithCode(7507, CODE)).toBe('7,507 ');
    expect(formatChartValWithCode(2208, CODE)).toBe('2,208 ');
    expect(formatChartValWithCode(6117, CODE)).toBe('6,117 ');
  });
  it('negative value uses the parenthesised section', () => {
    expect(formatChartValWithCode(-7507, CODE)).toBe('(7,507)');
  });
});

describe('formatChartValWithCode — percent & null', () => {
  it('scales by 100 with %', () => {
    expect(formatChartValWithCode(0.5, '0%')).toBe('50%');
  });
  it('falls back to General when code is null/General', () => {
    expect(formatChartValWithCode(7507, null)).toBe('7507');
  });
});
