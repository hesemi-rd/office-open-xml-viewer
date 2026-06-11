import { describe, it, expect } from 'vitest';
import { tableOverlayBorder, type TableCellStyle } from './renderer.js';
import type { Dxf } from './types.js';

/**
 * Regression tests for spurious table borders on *custom* `<tableStyle>`s
 * (ECMA-376 §18.8.83 / §18.5.1.2).
 *
 * A custom style contributes only the dxfs of its declared
 * `<tableStyleElement>`s. When none of those dxfs define a border, Excel draws
 * no table-level border at all — the only structure lines come from theme
 * borders baked into each cell `xf`. The renderer previously synthesized
 * accent-colored rules in that case (an approximation meant only for built-in
 * `TableStyle{Light,Medium,Dark}N` styles whose definitions are absent from the
 * file), which produced gray lines Excel never draws (sample-6 B23:E23 etc.).
 */

function cell(overrides: Partial<TableCellStyle>): TableCellStyle {
  return {
    accent: '#808080',
    isCustom: false,
    isHeader: false,
    isTotals: false,
    isBanded: false,
    isFirstCol: false,
    isLastCol: false,
    isTopEdge: false,
    isBottomEdge: false,
    ...overrides,
  };
}

const borderDxf: Dxf = {
  font: null,
  fill: null,
  border: { left: null, right: null, top: null, bottom: { style: 'thin', color: '#000000' } },
};

describe('tableOverlayBorder', () => {
  it('custom style with no border dxf draws nothing (no accent synthesis)', () => {
    // sample-6: custom "交通費" header row, dxf has fill only, no border.
    const ts = cell({ isCustom: true, isHeader: true, isTopEdge: true });
    const overlay = tableOverlayBorder(ts, undefined, undefined, 0);
    expect(overlay.kind).toBe('none');
  });

  it('custom style with no border dxf on a data row draws nothing', () => {
    const ts = cell({ isCustom: true, isBanded: true });
    const overlay = tableOverlayBorder(ts, undefined, undefined, 2);
    expect(overlay.kind).toBe('none');
  });

  it('custom style WITH a header-row border dxf draws that border', () => {
    const ts = cell({ isCustom: true, isHeader: true });
    const overlay = tableOverlayBorder(ts, undefined, borderDxf, 0);
    expect(overlay.kind).toBe('dxf');
    if (overlay.kind === 'dxf') {
      expect(overlay.border.bottom?.style).toBe('thin');
    }
  });

  it('built-in style with no border dxf still synthesizes accent rules', () => {
    // Built-in TableStyleLight* etc. are not in the file; the accent
    // approximation must remain until we ship a real preset catalog.
    const ts = cell({ isCustom: false, isHeader: true, isTopEdge: true, accent: '#4472C4' });
    const overlay = tableOverlayBorder(ts, undefined, undefined, 0);
    expect(overlay.kind).toBe('accent');
    if (overlay.kind === 'accent') {
      expect(overlay.color).toBe('#4472C4');
      expect(overlay.topEdge).toBe(true);
    }
  });
});
