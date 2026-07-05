import { describe, it, expect, vi } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import { XlsxWorkbook } from './workbook.js';

/**
 * `XlsxWorkbook.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight so a
 * `load()` / image extraction awaiting the worker cannot hang after the
 * workbook is disposed. Pinned with a real {@link WorkerBridge} over an
 * in-memory worker (the constructor opens a real Worker, so we build
 * off-prototype and inject the collaborators destroy() touches — the pattern
 * from `workbook.image.test.ts`).
 */

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminated = false;
  terminate(): void {
    this.terminated = true;
  }
}

interface DestroyProbe {
  destroy(): void;
}

describe('XlsxWorkbook.destroy() — rejects in-flight worker requests', () => {
  function makeWorkbook() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance.sheetCache = new Map();
    instance.imageCache = new Map();
    instance.imageBlobCache = new Map();
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { wb: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { wb, bridge, worker } = makeWorkbook();
    const inFlight = bridge.request((id) => ({ id }));
    wb.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice', () => {
    const { wb } = makeWorkbook();
    wb.destroy();
    expect(() => wb.destroy()).not.toThrow();
  });
});

/**
 * `destroy()` must close every cached `ImageBitmap` (GPU-backed) before
 * dropping `imageCache`, not just `.clear()` it — a bare `.clear()` drops the
 * last reference without releasing the GPU backing, leaking it until GC (which
 * is not guaranteed to run promptly for GPU-backed objects). See
 * `closeAndClearImageCache` in render-orchestrator.ts.
 */
describe('XlsxWorkbook.destroy() — closes cached ImageBitmaps (GPU-leak guard)', () => {
  function makeWorkbookWithImageCache(imageCache: Map<string, CanvasImageSource | null>) {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    instance.sheetCache = new Map();
    instance.imageCache = imageCache;
    instance.imageBlobCache = new Map();
    instance._fetchImage = () => Promise.resolve(new Blob());
    return instance as unknown as DestroyProbe;
  }

  it('calls .close() on each cached ImageBitmap and empties the cache', () => {
    const close1 = vi.fn();
    const close2 = vi.fn();
    const bmp1 = { close: close1 } as unknown as ImageBitmap;
    const bmp2 = { close: close2 } as unknown as ImageBitmap;
    const imageCache = new Map<string, CanvasImageSource | null>([
      ['xl/media/image1.png', bmp1],
      ['xl/media/image2.png', bmp2],
    ]);
    const wb = makeWorkbookWithImageCache(imageCache);

    wb.destroy();

    expect(close1).toHaveBeenCalledTimes(1);
    expect(close2).toHaveBeenCalledTimes(1);
    expect(imageCache.size).toBe(0);
  });

  it('skips a cached null (unsupported metafile) without throwing', () => {
    const imageCache = new Map<string, CanvasImageSource | null>([
      ['xl/media/diagram.emf', null],
    ]);
    const wb = makeWorkbookWithImageCache(imageCache);
    expect(() => wb.destroy()).not.toThrow();
    expect(imageCache.size).toBe(0);
  });

  it('is safe to destroy() twice — the second call does not re-close an already-closed bitmap', () => {
    const close = vi.fn();
    const bmp = { close } as unknown as ImageBitmap;
    const imageCache = new Map<string, CanvasImageSource | null>([['xl/media/image1.png', bmp]]);
    const wb = makeWorkbookWithImageCache(imageCache);

    wb.destroy();
    expect(() => wb.destroy()).not.toThrow();
    // The map is empty after the first destroy(), so the second pass has
    // nothing to iterate — close() is called exactly once total.
    expect(close).toHaveBeenCalledTimes(1);
  });
});
