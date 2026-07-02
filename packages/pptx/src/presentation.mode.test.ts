import { describe, it, expect } from 'vitest';
import { PptxPresentation } from './presentation';

/**
 * `PptxPresentation.mode` is a public fact (WS4 O1): an injected engine's render
 * mode decides the scroll viewer's render path (renderSlide vs
 * renderSlideToBitmap) so the viewer routes without error-probing (design §11 —
 * no silent mis-pathing).
 *
 * The constructor opens a real Worker, so we build the instance off-prototype and
 * inject only the private `_mode` field the getter reads.
 */
describe('PptxPresentation.mode', () => {
  const make = (mode: 'main' | 'worker') => {
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._mode = mode;
    return instance as unknown as PptxPresentation;
  };

  it('reflects the worker mode the engine was loaded with', () => {
    expect(make('worker').mode).toBe('worker');
  });

  it('reflects the main mode the engine was loaded with', () => {
    expect(make('main').mode).toBe('main');
  });
});
