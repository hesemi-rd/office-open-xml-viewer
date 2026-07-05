import { describe, it, expect, afterEach } from 'vitest';
import { WorkerBridge, registerEmbeddedFonts, type WorkerLike } from '@silurus/ooxml-core';
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

// ── Fake FontFaceSet so destroy()'s embedded-font release is observable ──────
const G = globalThis as Record<string, unknown>;
const ORIG_FONTS = { document: G.document, self: G.self, FontFace: G.FontFace };
afterEach(() => {
  G.document = ORIG_FONTS.document;
  G.self = ORIG_FONTS.self;
  G.FontFace = ORIG_FONTS.FontFace;
});

interface FakeFace { family: string }
function installFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: ArrayBuffer, public descriptors?: object) {}
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  const set = {
    faces: added,
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  delete G.self;
  return { added };
}

/** A minimal valid sfnt header so registerEmbeddedFonts accepts the face. */
const validHeader = (): Uint8Array =>
  new Uint8Array([
    0x00, 0x01, 0x00, 0x00, 0x00, 0x10, 0x01, 0x00, 0x00, 0x40, 0x00, 0x30,
    0x47, 0x53, 0x55, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  ]);

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
    instance._embeddedFontFaces = [];
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

  // Wiring guard: destroy() must actually release the embedded fonts the document
  // registered into the shared FontFaceSet. The other tests set
  // `_embeddedFontFaces = []`, so they never exercise the unregister branch — a
  // dropped call (or a wrong field name) would go unnoticed. Register a real face
  // through core, hand it to the document, then assert destroy() removes it from
  // the (fake) FontFaceSet and clears the held array.
  it('destroy() releases the document’s embedded fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await registerEmbeddedFonts([
      { family: 'DocxEmbedded', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(added).toHaveLength(1); // the face is in the shared set

    const { doc } = makeDocument();
    (doc as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces = held;
    doc.destroy();

    // destroy() called unregisterEmbeddedFonts(held): last holder gone → the face
    // left the FontFaceSet, and the held array was cleared.
    expect(added).toHaveLength(0);
    expect((doc as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces).toHaveLength(0);
  });
});
