import { describe, it, expect, afterEach, vi } from 'vitest';
import { preloadGoogleFonts, parseFontFaceRules, _resetCssCacheForTests, type FontPreloadEntry } from './preload.js';

const G = globalThis as Record<string, unknown>;
const ORIG = { document: G.document, self: G.self, fetch: G.fetch, FontFace: G.FontFace };

afterEach(() => {
  G.document = ORIG.document;
  G.self = ORIG.self;
  G.fetch = ORIG.fetch;
  G.FontFace = ORIG.FontFace;
  _resetCssCacheForTests();
  vi.restoreAllMocks();
});

const CSS = `
/* latin */
@font-face {
  font-family: 'Carlito';
  font-style: italic;
  font-weight: 700;
  src: url(https://fonts.gstatic.com/s/carlito/x.woff2) format('woff2');
  unicode-range: U+0000-00FF;
}
@font-face {
  font-family: 'Carlito';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2');
}
`;

describe('parseFontFaceRules', () => {
  it('extracts family, src and descriptors from @font-face blocks', () => {
    const faces = parseFontFaceRules(CSS);
    expect(faces).toHaveLength(2);
    expect(faces[0].family).toBe('Carlito');
    expect(faces[0].src).toContain('woff2');
    expect(faces[0].descriptors).toMatchObject({
      style: 'italic',
      weight: '700',
      unicodeRange: 'U+0000-00FF',
    });
    expect(faces[1].descriptors.style).toBe('normal');
  });
});

interface FakeFace {
  family: string;
  source: string;
  loadCalls: number;
  load: () => Promise<FakeFace>;
}

function installFakes(opts: { failLoad?: boolean } = {}) {
  const added: FakeFace[] = [];
  class FakeFontFace implements FakeFace {
    family: string; source: string; loadCalls = 0;
    constructor(family: string, source: string, public descriptors?: object) {
      this.family = family; this.source = source;
    }
    load(): Promise<FakeFace> {
      this.loadCalls++;
      return opts.failLoad ? Promise.reject(new Error('net')) : Promise.resolve(this);
    }
  }
  const set = {
    faces: added,
    add: (f: FakeFace) => { added.push(f); },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.fetch = vi.fn(async () => ({ ok: true, text: async () => CSS }));
  return { set, added };
}

const MAP: Record<string, FontPreloadEntry> = {
  calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
};

describe('preloadGoogleFonts', () => {
  it('fetches the CSS once, registers FontFaces and force-loads them (document.fonts)', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    await preloadGoogleFonts(['Calibri'], MAP);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
    expect(added.map((f) => f.family)).toEqual(['Carlito', 'Carlito']);
    expect(added.every((f) => f.loadCalls === 1)).toBe(true);
  });

  it('uses self.fonts when document is undefined (worker)', async () => {
    const { set, added } = installFakes();
    delete G.document;
    G.self = { fonts: set };
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(2);
  });

  it('no-ops without any FontFaceSet', async () => {
    delete G.document;
    delete G.self;
    await expect(preloadGoogleFonts(['Calibri'], MAP)).resolves.toBeUndefined();
  });

  it('does not refetch a CSS url already registered in this context', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    await preloadGoogleFonts(['Calibri'], MAP);
    await preloadGoogleFonts(['Calibri'], MAP);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
  });

  it('warns once when faces fail to load, and still resolves', async () => {
    const { set } = installFakes({ failLoad: true });
    G.document = { fonts: set };
    delete G.self;
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toContain('carlito');
  });

  it('warns once on a non-ok fetch and frees the slot so a later call retries', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    // Override fetch with a 404 so registration fails (res.ok === false).
    G.fetch = vi.fn(async () => ({ ok: false, status: 404, text: async () => '' }));
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toContain('carlito');
    // Failed url is removed from the cache, so a second call re-fetches.
    await preloadGoogleFonts(['Calibri'], MAP);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
  });

  it('joins a concurrent same-url call: second does not resolve before the first registers', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    // Defer the fetch so both calls overlap with the registration in flight.
    let resolveFetch!: (v: { ok: boolean; text: () => Promise<string> }) => void;
    const fetchPromise = new Promise<{ ok: boolean; text: () => Promise<string> }>((r) => {
      resolveFetch = r;
    });
    G.fetch = vi.fn(() => fetchPromise);

    let aDone = false;
    let bDone = false;
    const a = preloadGoogleFonts(['Calibri'], MAP).then(() => { aDone = true; });
    const b = preloadGoogleFonts(['Calibri'], MAP).then(() => { bDone = true; });

    // Let microtasks flush; neither call may settle while the fetch is pending.
    await Promise.resolve();
    await Promise.resolve();
    expect(aDone).toBe(false);
    expect(bDone).toBe(false);

    resolveFetch({ ok: true, text: async () => CSS });
    await Promise.all([a, b]);
    expect(aDone).toBe(true);
    expect(bDone).toBe(true);
    // Faces are registered exactly once (the joined call does not re-add them)
    // and the url is fetched once. Each face is force-loaded — both callers run
    // step 2 against the shared faces, so loadCalls is per-caller, but the
    // registration (add + fetch) happened a single time.
    expect(added.map((f) => f.family)).toEqual(['Carlito', 'Carlito']);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
    expect(added.every((f) => f.loadCalls >= 1)).toBe(true);
  });
});
