import { describe, it, expect } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument } from './document';

/**
 * `DocxDocument.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight —
 * otherwise a `load()` / image extraction awaiting the worker would hang
 * forever after the document is disposed. This pins that delegation using a
 * real {@link WorkerBridge} over an in-memory worker (no real Worker: the
 * constructor opens one, so we build the instance off-prototype and inject the
 * bridge — the established pattern from `document.image.test.ts`).
 */

/** In-memory Worker stand-in that never answers, so requests stay pending until
 *  the bridge is terminated. */
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

describe('DocxDocument.destroy() — rejects in-flight worker requests', () => {
  function makeDocument() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
    instance._bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance._imageCache = new Map();
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { doc: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { doc, bridge, worker } = makeDocument();
    // A request the worker will never answer.
    const inFlight = bridge.request((id) => ({ id }));
    doc.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice (second terminate has nothing pending)', () => {
    const { doc } = makeDocument();
    doc.destroy();
    expect(() => doc.destroy()).not.toThrow();
  });
});
