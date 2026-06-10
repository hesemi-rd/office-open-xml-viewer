import { describe, it, expect, afterEach, vi } from 'vitest';
import { preloadGoogleFonts, type FontPreloadEntry } from './preload.js';

// Minimal DOM/FontFaceSet doubles so the loader's promise ordering can be
// asserted in the node test environment (no jsdom). We model the two failure
// modes the real flicker came from: (1) the stylesheet <link> registering its
// FontFace entries asynchronously, and (2) face.load() settling later than the
// loader used to wait. The loader must not resolve until BOTH have settled.

interface FakeFace {
  family: string;
  loaded: boolean;
  resolveLoad: () => void;
  load: () => Promise<void>;
}

function makeFace(family: string): FakeFace {
  let resolve!: () => void;
  const p = new Promise<void>((r) => (resolve = r));
  const face: FakeFace = {
    family,
    loaded: false,
    resolveLoad: () => {
      face.loaded = true;
      resolve();
    },
    load: () => p,
  };
  return face;
}

const ORIG = {
  document: (globalThis as Record<string, unknown>).document,
};

afterEach(() => {
  (globalThis as Record<string, unknown>).document = ORIG.document;
  vi.restoreAllMocks();
});

const MAP: Record<string, FontPreloadEntry> = {
  carlito: { url: 'https://fonts.example/carlito.css', loadFamily: 'Carlito' },
};

function installFakeDom(faces: FakeFace[]) {
  const linkListeners: Record<string, (() => void)[]> = {};
  let appendedLink: Record<string, unknown> | null = null;

  const link: Record<string, unknown> = {
    rel: '',
    href: '',
    sheet: null, // not yet parsed
    addEventListener: (ev: string, cb: () => void) => {
      (linkListeners[ev] ??= []).push(cb);
    },
    removeEventListener: () => {},
  };

  const doc = {
    querySelector: () => null,
    createElement: () => link,
    head: { appendChild: (l: Record<string, unknown>) => (appendedLink = l) },
    fonts: faces as unknown as Iterable<FakeFace> & { ready: Promise<unknown> },
  };
  (doc.fonts as unknown as { ready: Promise<unknown> }).ready = Promise.resolve();
  (globalThis as Record<string, unknown>).document = doc;

  return {
    fireLinkLoad: () => {
      (linkListeners['load'] ?? []).forEach((cb) => cb());
    },
    get appendedLink() {
      return appendedLink;
    },
  };
}

describe('preloadGoogleFonts', () => {
  it('does not resolve until the stylesheet loads AND every face.load() settles', async () => {
    const face = makeFace('Carlito');
    const dom = installFakeDom([face]);

    let settled = false;
    const p = preloadGoogleFonts(['Calibri'], { calibri: MAP.carlito }).then(() => {
      settled = true;
    });

    // Give microtasks a chance; must still be pending (link not loaded yet).
    await Promise.resolve();
    expect(settled).toBe(false);

    // Stylesheet parses → FontFace entries registered, but face still loading.
    dom.fireLinkLoad();
    await Promise.resolve();
    await Promise.resolve();
    expect(settled).toBe(false);
    expect(face.loaded).toBe(false);

    // Face finishes downloading → loader resolves.
    face.resolveLoad();
    await p;
    expect(settled).toBe(true);
    expect(face.loaded).toBe(true);
  });

  it('returns immediately when no requested family is in the map', async () => {
    installFakeDom([]);
    await expect(
      preloadGoogleFonts(['Totally Unknown Face'], MAP),
    ).resolves.toBeUndefined();
  });

  it('is a no-op without a document (SSR / worker)', async () => {
    (globalThis as Record<string, unknown>).document = undefined;
    await expect(preloadGoogleFonts(['Calibri'], MAP)).resolves.toBeUndefined();
  });
});
