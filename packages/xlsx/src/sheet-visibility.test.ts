import { describe, it, expect } from 'vitest';
import { selectSheetVisibility } from './sheet-visibility.js';
import type { SheetMeta } from './types.js';

/** Build a sheet list where `vis[i]` sets SheetMeta.visibility. */
function sheets(vis: Array<'hidden' | 'veryHidden' | undefined>): SheetMeta[] {
  return vis.map((v, i) => ({
    name: `S${i}`,
    sheetId: i + 1,
    rId: `rId${i + 1}`,
    ...(v ? { visibility: v } : {}),
  }));
}

describe('selectSheetVisibility (XlsxWorkbook.sheetVisibility core)', () => {
  it('returns the visibility for an in-range sheet', () => {
    const s = sheets([undefined, 'hidden', 'veryHidden']);
    expect(selectSheetVisibility(s, 0)).toBe('visible');
    expect(selectSheetVisibility(s, 1)).toBe('hidden');
    expect(selectSheetVisibility(s, 2)).toBe('veryHidden');
  });
  it('treats absent visibility as visible', () => {
    expect(selectSheetVisibility(sheets([undefined]), 0)).toBe('visible');
  });
  it('returns visible for out-of-range / non-integer (non-clamped, like pptx)', () => {
    const s = sheets(['hidden']);
    expect(selectSheetVisibility(s, -1)).toBe('visible');
    expect(selectSheetVisibility(s, 1)).toBe('visible');
    expect(selectSheetVisibility(s, 0.5)).toBe('visible');
    expect(selectSheetVisibility([], 0)).toBe('visible');
  });
});
