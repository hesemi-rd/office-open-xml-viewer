import { describe, it, expect, afterEach } from 'vitest';
import {
  WorkerBridge,
  loadLocalFontMetrics,
  registerEmbeddedFonts,
  preloadGoogleFonts,
  type WorkerLike,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';

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

// ── Fake FontFaceSet so destroy()'s embedded-font / Google-Fonts release is
// observable ──────────────────────────────────────────────────────────────
const G = globalThis as Record<string, unknown>;
const ORIG_FONTS = {
  document: G.document,
  self: G.self,
  fetch: G.fetch,
  FontFace: G.FontFace,
  OffscreenCanvas: G.OffscreenCanvas,
};
afterEach(() => {
  G.document = ORIG_FONTS.document;
  G.self = ORIG_FONTS.self;
  G.fetch = ORIG_FONTS.fetch;
  G.FontFace = ORIG_FONTS.FontFace;
  G.OffscreenCanvas = ORIG_FONTS.OffscreenCanvas;
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

// ── Google-Fonts flavored fake: `preloadGoogleFonts` needs `fetch` (to pull
// the CSS) and a string-`src` `FontFace` constructor, unlike the ArrayBuffer
// source used by the embedded-font fake above. Mirrors the fake used by
// `presentation-destroy.test.ts` / `workbook-destroy.test.ts`. ──────────────
const GOOGLE_CSS = `@font-face { font-family: 'Carlito'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2'); }`;
function installGoogleFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: string, public descriptors?: object) {}
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  const set = {
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  G.fetch = async () => ({ ok: true, text: async () => GOOGLE_CSS });
  delete G.self;
  return { added };
}

function installLocalMetricFontEnvironment(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    status: FontFaceLoadStatus = 'unloaded';
    constructor(public family: string, public source: string) {}
    load(): Promise<this> {
      this.status = 'loaded';
      return Promise.resolve(this);
    }
  }
  const set = {
    add: (face: FakeFace) => { added.push(face); },
    delete: (face: FakeFace) => {
      const index = added.indexOf(face);
      if (index >= 0) added.splice(index, 1);
      return index >= 0;
    },
  };
  class FakeOffscreenCanvas {
    getContext() {
      return {
        font: '',
        measureText: () => ({ fontBoundingBoxAscent: 106, fontBoundingBoxDescent: 44 }),
      };
    }
  }
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  G.OffscreenCanvas = FakeOffscreenCanvas;
  delete G.self;
  return { added };
}
const GOOGLE_FONT_MAP: Record<string, FontPreloadEntry> = {
  calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
};

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
    attachDocumentLayoutRuntime(instance, 0);
    instance._bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance._imageCache = new Map();
    instance._embeddedFontFaces = [];
    instance._googleFontFaces = [];
    instance._localMetricFontFaces = [];
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

  // Wiring guard: destroy() must actually release the Google-Fonts substitutes
  // the document preloaded into the shared FontFaceSet. The other tests set
  // `_googleFontFaces = []`, so they never exercise the unload branch — a
  // dropped call (or a wrong field name) would go unnoticed. Preload a real
  // face through core, hand it to the document, then assert destroy() removes
  // it from the (fake) FontFaceSet and clears the held array. Twin of the
  // embedded-fonts guard above; same shape as
  // `presentation-destroy.test.ts` / `workbook-destroy.test.ts`.
  it('destroy() releases the document’s Google fonts from the FontFaceSet', async () => {
    const { added } = installGoogleFontFaceSet();
    const held = await preloadGoogleFonts(['Calibri'], GOOGLE_FONT_MAP);
    expect(added).toHaveLength(1); // the web font is in the shared set

    const { doc } = makeDocument();
    (doc as unknown as { _googleFontFaces: FontFace[] })._googleFontFaces = held;
    doc.destroy();

    // destroy() called unloadGoogleFonts(held): last holder gone → the face
    // left the FontFaceSet, and the held array was cleared.
    expect(added).toHaveLength(0);
    expect((doc as unknown as { _googleFontFaces: FontFace[] })._googleFontFaces).toHaveLength(0);
  });

  it('destroy() releases exact local metric faces from the FontFaceSet', async () => {
    const { added } = installLocalMetricFontEnvironment();
    const held = await loadLocalFontMetrics([{
      family: 'Metric Face',
      localNames: ['Metric Face'],
      lineHeightMultiplier: 1.3,
    }]);
    expect(added).toHaveLength(1);

    const { doc } = makeDocument();
    (doc as unknown as { _localMetricFontFaces: FontFace[] })._localMetricFontFaces = held.faces;
    doc.destroy();

    expect(added).toHaveLength(0);
    expect((doc as unknown as { _localMetricFontFaces: FontFace[] })._localMetricFontFaces).toHaveLength(0);
  });
});
