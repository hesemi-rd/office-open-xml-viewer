import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  verticalVertFeatureSupported,
  withVertFeature,
} from './vertical-vert-feature.js';

type Listener = () => void;

interface ProbeFixture {
  ctx: CanvasRenderingContext2D;
  featuresAtPaint: string[];
  fontAssignments: () => number;
  paints: () => number;
  listeners: Map<string, Listener>;
  supported: Set<number>;
}

function fakeHtmlCanvasProbe(initialFeature = 'normal'): ProbeFixture {
  const listeners = new Map<string, Listener>();
  const supported = new Set<number>();
  const featuresAtPaint: string[] = [];
  let paints = 0;
  let fontAssignments = 0;

  class FakeCanvas {
    width = 1;
    height = 1;
    isConnected = false;
    ownerDocument: typeof document;
    style = { fontFeatureSettings: initialFeature };
    private context: CanvasRenderingContext2D | null = null;

    constructor(ownerDocument: typeof document) {
      this.ownerDocument = ownerDocument;
    }

    setAttribute() {}
    remove() { this.isConnected = false; }

    getContext(): CanvasRenderingContext2D {
      if (this.context !== null) return this.context;
      const canvas = this;
      let font = '';
      let resolvedFeature = canvas.style.fontFeatureSettings;
      let drawnCp = 0;
      this.context = {
        canvas,
        get font() { return font; },
        set font(value: string) {
          font = value;
          resolvedFeature = canvas.style.fontFeatureSettings;
          fontAssignments += 1;
        },
        fillStyle: '#000',
        textAlign: 'center',
        textBaseline: 'middle',
        clearRect() {},
        fillText(text: string) {
          drawnCp = text.codePointAt(0) ?? 0;
          featuresAtPaint.push(resolvedFeature);
          paints += 1;
        },
        getImageData() {
          const data = new Uint8ClampedArray(canvas.width * canvas.height * 4);
          const vert = resolvedFeature.includes('"vert" 1') && supported.has(drawnCp);
          const inkWidth = vert ? 3 : 9;
          const inkHeight = vert ? 9 : 3;
          for (let y = 0; y < inkHeight; y += 1) {
            for (let x = 0; x < inkWidth; x += 1) {
              data[(y * canvas.width + x) * 4 + 3] = 255;
            }
          }
          return { data } as ImageData;
        },
      } as unknown as CanvasRenderingContext2D;
      return this.context;
    }
  }

  const parent = {
    appendChild(canvas: FakeCanvas) { canvas.isConnected = true; },
  };
  const fakeDocument = {
    body: parent,
    documentElement: parent,
    fonts: {
      addEventListener(name: string, listener: Listener) { listeners.set(name, listener); },
    },
    createElement: () => new FakeCanvas(fakeDocument as unknown as Document),
  } as unknown as Document;
  vi.stubGlobal('HTMLCanvasElement', FakeCanvas);
  vi.stubGlobal('document', fakeDocument);

  const canvas = new FakeCanvas(fakeDocument);
  let sourceFont = '16px "Probe Mincho", serif';
  const ctx = {
    canvas,
    get font() { return sourceFont; },
    set font(value: string) { sourceFont = value; },
  } as unknown as CanvasRenderingContext2D;
  return {
    ctx,
    featuresAtPaint,
    fontAssignments: () => fontAssignments,
    paints: () => paints,
    listeners,
    supported,
  };
}

describe('vertical OpenType vert feature', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reports unsupported for a context whose canvas has no element style', () => {
    const ctx = {
      canvas: { width: 1, height: 1 },
      font: '16px serif',
    } as unknown as OffscreenCanvasRenderingContext2D;

    expect(verticalVertFeatureSupported(ctx, 0x30fc)).toBe(false);
  });

  it('probes the requested glyph, refonts after each feature change, and caches per code point', () => {
    const fixture = fakeHtmlCanvasProbe('"kern" 1, "ss01" 1');
    fixture.supported.add(0x30fc);

    expect(verticalVertFeatureSupported(fixture.ctx, 0x30fc)).toBe(true);
    expect(fixture.featuresAtPaint).toEqual([
      '"kern" 1, "ss01" 1',
      '"kern" 1, "ss01" 1, "vert" 1',
    ]);
    expect(fixture.fontAssignments()).toBeGreaterThanOrEqual(3);
    expect(verticalVertFeatureSupported(fixture.ctx, 0x30fc)).toBe(true);
    expect(fixture.paints()).toBe(2);

    expect(verticalVertFeatureSupported(fixture.ctx, 0x301c)).toBe(false);
    expect(fixture.paints()).toBe(4);
  });

  it('invalidates each glyph result after the FontFace epoch changes', () => {
    const fixture = fakeHtmlCanvasProbe();

    expect(verticalVertFeatureSupported(fixture.ctx, 0xff5e)).toBe(false);
    fixture.supported.add(0xff5e);
    expect(verticalVertFeatureSupported(fixture.ctx, 0xff5e)).toBe(false);
    expect(fixture.paints()).toBe(2);

    fixture.listeners.get('loadingdone')?.();
    expect(verticalVertFeatureSupported(fixture.ctx, 0xff5e)).toBe(true);
    expect(fixture.paints()).toBe(4);
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
      'feature:"kern" 1, "vert" 1',
      'font:italic 700 16px "Hiragino Mincho ProN"',
      'draw:"kern" 1, "vert" 1',
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
