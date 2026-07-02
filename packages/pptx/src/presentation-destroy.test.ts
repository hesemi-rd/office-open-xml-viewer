import { describe, it, expect } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import { PptxPresentation } from './presentation';

/**
 * `PptxPresentation.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight so a
 * `load()` / image extraction awaiting the worker cannot hang after the deck is
 * disposed. Pinned with a real {@link WorkerBridge} over an in-memory worker
 * (the constructor opens a real Worker, so we build off-prototype and inject
 * the collaborators destroy() touches — the pattern from
 * `presentation.image.test.ts`).
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

describe('PptxPresentation.destroy() — rejects in-flight worker requests', () => {
  function makePresentation() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance._mediaCache = new Map();
    instance._imageCache = new Map();
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { pres: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { pres, bridge, worker } = makePresentation();
    const inFlight = bridge.request((id) => ({ id }));
    pres.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice', () => {
    const { pres } = makePresentation();
    pres.destroy();
    expect(() => pres.destroy()).not.toThrow();
  });
});
