import { describe, it, expect } from 'vitest';
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
