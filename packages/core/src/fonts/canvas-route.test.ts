import { describe, expect, it } from 'vitest';
import { canvasFontString, createCanvasFontRoute } from './canvas-route.js';

describe('Canvas font route', () => {
  it('serializes the immutable complete family list without adding a fallback', () => {
    const route = createCanvasFontRoute('"Times New Roman", serif', 'native');
    expect(canvasFontString(route, 12, 700, 'italic'))
      .toBe('italic 700 12px "Times New Roman", serif');
    expect(createCanvasFontRoute('"Times New Roman", serif', 'native').fingerprint)
      .toBe(route.fingerprint);
    expect(createCanvasFontRoute('"Times New Roman", serif', 'registered').fingerprint)
      .not.toBe(route.fingerprint);
    expect(createCanvasFontRoute('a:b', 'native').fingerprint)
      .not.toBe(createCanvasFontRoute('a%3Ab', 'native').fingerprint);
  });
});
