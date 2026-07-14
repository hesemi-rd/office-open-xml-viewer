import { describe, expect, it } from 'vitest';
import { layoutOptionsKey, normalizeLayoutOptions } from './options.js';
import type { LayoutServices } from './types.js';

function services(text: string, images: string, math: string): LayoutServices {
  return {
    text: { fingerprint: text, localMetrics: {}, shape: () => ({ advancePt: 0, ascentPt: 0, descentPt: 0, spans: [], diagnostics: [] }) },
    images: { fingerprint: images, resolve: () => ({ widthPt: 1, heightPt: 1, mimeType: 'image/png' }) },
    math: { fingerprint: math, resolve: () => ({ resourceKey: 'm', widthEm: 1, ascentEm: 1, descentEm: 0, diagnostics: [] }) },
  };
}

describe('layout options', () => {
  it('normalizes Date, number, and undefined against one captured load-time default', () => {
    expect(normalizeLayoutOptions(new Date(123), 999)).toEqual({ currentDateMs: 123 });
    expect(normalizeLayoutOptions(456, 999)).toEqual({ currentDateMs: 456 });
    expect(normalizeLayoutOptions(undefined, 999)).toEqual({ currentDateMs: 999 });

    if (false) {
      // @ts-expect-error environment strings are not a layout input
      normalizeLayoutOptions(undefined, 'browser-fonts-v1');
    }
  });

  it('keys only the normalized date and actual service fingerprints', () => {
    const base = services('text:a', 'images:a', 'math:a');
    const key = layoutOptionsKey({ currentDateMs: 100 }, base);

    expect(layoutOptionsKey({ currentDateMs: 101 }, base)).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('text:b', 'images:a', 'math:a'))).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('text:a', 'images:b', 'math:a'))).not.toBe(key);
    expect(layoutOptionsKey({ currentDateMs: 100 }, services('text:a', 'images:a', 'math:b'))).not.toBe(key);

    if (false) {
      // @ts-expect-error paint width, DPR, and color cannot enter the layout key
      layoutOptionsKey({ currentDateMs: 100 }, base, { width: 600, dpr: 2, defaultTextColor: '#fff' });
    }
  });
});
