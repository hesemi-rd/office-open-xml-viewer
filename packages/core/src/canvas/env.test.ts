import { describe, it, expect, afterEach } from 'vitest';
import { isHTMLCanvas, defaultDpr } from './env.js';

const G = globalThis as Record<string, unknown>;
const ORIG = { HTMLCanvasElement: G.HTMLCanvasElement, window: G.window };

afterEach(() => {
  G.HTMLCanvasElement = ORIG.HTMLCanvasElement;
  G.window = ORIG.window;
});

describe('isHTMLCanvas', () => {
  it('returns false without throwing when HTMLCanvasElement is undefined (worker)', () => {
    delete G.HTMLCanvasElement;
    expect(isHTMLCanvas({ width: 1, height: 1 })).toBe(false);
  });

  it('detects an instance when the global exists', () => {
    class FakeCanvas {}
    G.HTMLCanvasElement = FakeCanvas;
    expect(isHTMLCanvas(new FakeCanvas())).toBe(true);
    expect(isHTMLCanvas({})).toBe(false);
  });
});

describe('defaultDpr', () => {
  it('falls back when window is undefined (worker)', () => {
    delete G.window;
    expect(defaultDpr()).toBe(1);
    expect(defaultDpr(2)).toBe(2);
  });

  it('reads window.devicePixelRatio on the main thread', () => {
    G.window = { devicePixelRatio: 2 };
    expect(defaultDpr()).toBe(2);
  });

  it('falls back when devicePixelRatio is 0/undefined', () => {
    G.window = {};
    expect(defaultDpr()).toBe(1);
  });
});
