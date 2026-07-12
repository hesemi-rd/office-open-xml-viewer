import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  verticalFallbackShearCoefficient,
  verticalFallbackShearEnabled,
} from './vertical-fallback-shear.js';

type Listener = () => void;

function rasterCanvas(slope: number, onRead?: () => void) {
  return class RasterCanvas {
    width: number;
    height: number;
    constructor(width: number, height: number) {
      this.width = width;
      this.height = height;
    }
    getContext(): CanvasRenderingContext2D {
      const canvas = this;
      return {
        canvas,
        font: '',
        fillStyle: '#000',
        textAlign: 'center',
        textBaseline: 'middle',
        clearRect() {},
        fillText() {},
        getImageData() {
          onRead?.();
          const data = new Uint8ClampedArray(canvas.width * canvas.height * 4);
          const x0 = Math.floor(canvas.width / 4);
          const x1 = Math.floor(canvas.width * 3 / 4);
          for (let x = x0; x <= x1; x++) {
            const center = canvas.height / 2 + slope * (x - canvas.width / 2);
            const y = Math.floor(center);
            const fraction = center - y;
            data[(y * canvas.width + x) * 4 + 3] = Math.round(255 * (1 - fraction));
            data[((y + 1) * canvas.width + x) * 4 + 3] = Math.round(255 * fraction);
          }
          // A damaged terminal column must not steer the robust all-column fit.
          data[((canvas.height - 2) * canvas.width + x0) * 4 + 3] = 255;
          return { data } as ImageData;
        },
      } as unknown as CanvasRenderingContext2D;
    }
  };
}

describe('non-vert prolonged-mark fallback shear', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('classifies U+30FC only, leaving both wave marks unchanged', () => {
    expect(verticalFallbackShearEnabled(0x30fc)).toBe(true);
    expect(verticalFallbackShearEnabled(0x301c)).toBe(false);
    expect(verticalFallbackShearEnabled(0xff5e)).toBe(false);
    expect(verticalFallbackShearEnabled(0x3030)).toBe(false);
  });

  it('measures an all-column alpha-centroid slope with robust Theil-Sen regression', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    const Canvas = rasterCanvas(0.25);
    const ctx = {
      canvas: new Canvas(10, 10),
      font: 'italic 600 17px "Probe Mincho", serif',
    } as unknown as CanvasRenderingContext2D;

    expect(verticalFallbackShearCoefficient(ctx, 0x30fc)).toBeCloseTo(0.25, 2);
  });

  it('returns zero without allocating for wave marks and on readback failure', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    let constructions = 0;
    class BrokenCanvas {
      constructor(_width: number, _height: number) { constructions += 1; }
      getContext() { throw new Error('readback unavailable'); }
    }
    const ctx = {
      canvas: Object.create(BrokenCanvas.prototype),
      font: '16px Failure Mincho',
    } as unknown as CanvasRenderingContext2D;

    expect(verticalFallbackShearCoefficient(ctx, 0x301c)).toBe(0);
    expect(constructions).toBe(0);
    expect(verticalFallbackShearCoefficient(ctx, 0x30fc)).toBe(0);
    expect(constructions).toBe(1);
  });

  it('returns zero when the raster has no ink or fewer than two ink columns', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    function coefficient(columns: number): number {
      class SparseCanvas {
        constructor(public width: number, public height: number) {}
        getContext() {
          const canvas = this;
          return {
            canvas, font: '', fillStyle: '#000', textAlign: 'center', textBaseline: 'middle',
            clearRect() {}, fillText() {},
            getImageData() {
              const data = new Uint8ClampedArray(canvas.width * canvas.height * 4);
              for (let x = 0; x < columns; x++) data[((256 * canvas.width + x) * 4) + 3] = 255;
              return { data };
            },
          };
        }
      }
      const ctx = { canvas: new SparseCanvas(1, 1), font: `16px Sparse${columns}` } as unknown as CanvasRenderingContext2D;
      return verticalFallbackShearCoefficient(ctx, 0x30fc);
    }

    expect(coefficient(0)).toBe(0);
    expect(coefficient(1)).toBe(0);
  });

  it('caches by size-independent font identity and remeasures after a FontFace epoch', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    const listeners = new Map<string, Listener>();
    let reads = 0;
    const Canvas = rasterCanvas(0.125, () => { reads += 1; });
    const fonts = {
      addEventListener: (name: string, listener: Listener) => listeners.set(name, listener),
    };
    const document = {
      fonts,
      createElement: () => new Canvas(1, 1),
    };
    vi.stubGlobal('document', document);
    const ctx = {
      canvas: { ownerDocument: document },
      font: 'normal 400 16px "Epoch Mincho", serif',
    } as unknown as CanvasRenderingContext2D;

    expect(verticalFallbackShearCoefficient(ctx, 0x30fc)).toBeCloseTo(0.125, 2);
    ctx.font = 'normal 400 48px "Epoch Mincho", serif';
    expect(verticalFallbackShearCoefficient(ctx, 0x30fc)).toBeCloseTo(0.125, 2);
    expect(reads).toBe(1);
    listeners.get('loadingdone')?.();
    expect(verticalFallbackShearCoefficient(ctx, 0x30fc)).toBeCloseTo(0.125, 2);
    expect(reads).toBe(2);
  });

  it('scopes non-DOM cache reuse to the source canvas surface', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    let reads = 0;
    const Canvas = rasterCanvas(0.125, () => { reads += 1; });
    const first = {
      canvas: new Canvas(10, 10),
      font: '16px "Surface Mincho", serif',
    } as unknown as CanvasRenderingContext2D;
    const second = {
      canvas: new Canvas(10, 10),
      font: '16px "Surface Mincho", serif',
    } as unknown as CanvasRenderingContext2D;

    expect(verticalFallbackShearCoefficient(first, 0x30fc)).toBeCloseTo(0.125, 2);
    expect(verticalFallbackShearCoefficient(first, 0x30fc)).toBeCloseTo(0.125, 2);
    expect(reads).toBe(1);
    expect(verticalFallbackShearCoefficient(second, 0x30fc)).toBeCloseTo(0.125, 2);
    expect(reads).toBe(2);
  });
});
