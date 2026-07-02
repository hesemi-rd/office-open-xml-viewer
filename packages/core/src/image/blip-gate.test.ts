import { describe, expect, it } from 'vitest';
import { preferVectorBlip } from './blip-gate.js';

describe('preferVectorBlip', () => {
  it('prefers the vector original when an svgImagePath is present and uncropped', () => {
    expect(preferVectorBlip({ svgImagePath: 'ppt/media/image1.svg' })).toBe(true);
    expect(preferVectorBlip({ svgImagePath: 'a.svg', srcRect: null })).toBe(true);
    expect(preferVectorBlip({ svgImagePath: 'a.svg', srcRect: undefined })).toBe(true);
  });

  it('falls back to the raster when there is no svgImagePath', () => {
    expect(preferVectorBlip({})).toBe(false);
    expect(preferVectorBlip({ svgImagePath: null })).toBe(false);
    expect(preferVectorBlip({ svgImagePath: undefined })).toBe(false);
  });

  it('falls back to the raster when a crop is present (any non-nullish srcRect)', () => {
    // The §20.1.8.55 crop math needs the raster's native pixel grid.
    expect(preferVectorBlip({ svgImagePath: 'a.svg', srcRect: { l: 0.1, t: 0, r: 0, b: 0 } })).toBe(
      false,
    );
    // Callers holding only an aggregated boolean crop flag pass `flag || null`.
    expect(preferVectorBlip({ svgImagePath: 'a.svg', srcRect: true })).toBe(false);
  });
});
