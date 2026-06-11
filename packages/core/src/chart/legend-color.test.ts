import { describe, it, expect } from 'vitest';
import type { ChartSeries } from '../types/chart';
import { legendEntryColor, CHART_PALETTE } from './renderer.js';

function series(partial: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...partial };
}

const pal = (i: number) => `#${CHART_PALETTE[i % CHART_PALETTE.length]}`;

describe('legendEntryColor', () => {
  describe('multi-series chart (bar/line) — one legend entry per series', () => {
    const s: ChartSeries[] = [
      series({ name: 'A', color: null }),
      series({ name: 'B', color: null }),
      series({ name: 'C', color: null }),
    ];

    it('falls back to the per-series palette index', () => {
      expect(legendEntryColor('bar', s, 0)).toBe(pal(0));
      expect(legendEntryColor('bar', s, 1)).toBe(pal(1));
      expect(legendEntryColor('bar', s, 2)).toBe(pal(2));
    });

    it('honors an explicit series color', () => {
      const withColor: ChartSeries[] = [
        series({ name: 'A', color: 'FF0000' }),
        series({ name: 'B', color: null }),
      ];
      expect(legendEntryColor('line', withColor, 0)).toBe('#FF0000');
      expect(legendEntryColor('line', withColor, 1)).toBe(pal(1));
    });
  });

  describe('pie / doughnut — one legend entry per category (data point of series[0])', () => {
    it('uses the per-index palette and matches slice colors, ignoring the series-level color', () => {
      // Excel sets a series-level solidFill on pie series; that must NOT
      // collapse every legend swatch to the same color. Each entry uses the
      // same palette index its slice does.
      const s: ChartSeries[] = [
        series({ name: 'Region', color: '4472C4', values: [10, 20, 30] }),
      ];
      expect(legendEntryColor('pie', s, 0)).toBe(pal(0));
      expect(legendEntryColor('pie', s, 1)).toBe(pal(1));
      expect(legendEntryColor('pie', s, 2)).toBe(pal(2));
      // Distinct colors, not all the series color.
      expect(legendEntryColor('pie', s, 0)).not.toBe(legendEntryColor('pie', s, 1));
    });

    it('honors explicit per-point dPt colors', () => {
      const s: ChartSeries[] = [
        series({
          name: 'Region',
          color: '4472C4',
          values: [10, 20, 30],
          dataPointColors: ['AA0000', null, 'CC0000'],
        }),
      ];
      expect(legendEntryColor('doughnut', s, 0)).toBe('#AA0000');
      // null inside the array -> palette fallback for that slice
      expect(legendEntryColor('doughnut', s, 1)).toBe(pal(1));
      expect(legendEntryColor('doughnut', s, 2)).toBe('#CC0000');
    });
  });
});
