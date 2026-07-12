import { describe, expect, it } from 'vitest';
import {
  verticalVertFeatureSupported,
  withVertFeature,
} from './vertical-vert-feature.js';

describe('vertical OpenType vert feature', () => {
  it('reports unsupported for a context whose canvas has no element style', () => {
    const ctx = {
      canvas: { width: 1, height: 1 },
      font: '16px serif',
    } as unknown as OffscreenCanvasRenderingContext2D;

    expect(verticalVertFeatureSupported(ctx)).toBe(false);
  });

  it('sets the canvas feature before refonting and restores both after drawing', () => {
    const events: string[] = [];
    let font = 'italic 700 16px "Hiragino Mincho ProN"';
    let feature = '"kern" 1';
    const style = {
      get fontFeatureSettings() {
        return feature;
      },
      set fontFeatureSettings(value: string) {
        feature = value;
        events.push(`feature:${value}`);
      },
    };
    const ctx = {
      canvas: { style },
      get font() {
        return font;
      },
      set font(value: string) {
        font = value;
        events.push(`font:${value}`);
      },
    } as unknown as CanvasRenderingContext2D;

    const result = withVertFeature(ctx, () => {
      events.push(`draw:${style.fontFeatureSettings}`);
      return 47;
    });

    expect(result).toBe(47);
    expect(events).toEqual([
      'feature:"vert" 1',
      'font:italic 700 16px "Hiragino Mincho ProN"',
      'draw:"vert" 1',
      'feature:"kern" 1',
      'font:italic 700 16px "Hiragino Mincho ProN"',
    ]);
  });

  it('restores the prior feature and refonts when drawing throws', () => {
    let fontAssignments = 0;
    let font = '16px serif';
    const style = { fontFeatureSettings: '' };
    const ctx = {
      canvas: { style },
      get font() {
        return font;
      },
      set font(value: string) {
        font = value;
        fontAssignments += 1;
      },
    } as unknown as CanvasRenderingContext2D;

    expect(() =>
      withVertFeature(ctx, () => {
        throw new Error('paint failed');
      }),
    ).toThrow('paint failed');
    expect(style.fontFeatureSettings).toBe('');
    expect(fontAssignments).toBe(2);
  });
});
