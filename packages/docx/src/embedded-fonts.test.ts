import { afterEach, describe, expect, it, vi } from 'vitest';
import { deobfuscateOdttf } from '@silurus/ooxml-core';
import { loadEmbeddedFonts } from './embedded-fonts.js';
import type { DocxDocumentModel, EmbeddedFontRef } from './types';

// `loadEmbeddedFonts` maps `doc.embeddedFonts` → `EmbeddedFontFace[]` and calls
// the real core `registerEmbeddedFonts`. We stub the global FontFace +
// FontFaceSet so the faces the mapper produces surface as `added` entries on a
// fake set — asserting the derived family / weight / style / odttf-plaintext,
// exactly as core's own embedded.test.ts does.

const G = globalThis as Record<string, unknown>;
const ORIG = { document: G.document, self: G.self, FontFace: G.FontFace };

afterEach(() => {
  G.document = ORIG.document;
  G.self = ORIG.self;
  G.FontFace = ORIG.FontFace;
  vi.restoreAllMocks();
});

interface FakeFace {
  family: string;
  source: ArrayBuffer;
  descriptors?: { weight?: string; style?: string };
  load: () => Promise<FakeFace>;
}

function installFontFaceSet() {
  const added: FakeFace[] = [];
  class FakeFontFace implements FakeFace {
    family: string;
    source: ArrayBuffer;
    constructor(
      family: string,
      source: ArrayBuffer,
      public descriptors?: { weight?: string; style?: string },
    ) {
      this.family = family;
      this.source = source;
    }
    load(): Promise<FakeFace> {
      return Promise.resolve(this);
    }
  }
  const set = {
    add: (f: FakeFace) => {
      added.push(f);
    },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  delete G.self;
  return added;
}

// A minimal, valid sfnt header (TrueType 0x00010000) so `FontFace(source)`
// would accept the bytes — mirrors core's `validHeader`.
const validHeader = () =>
  new Uint8Array([
    0x00, 0x01, 0x00, 0x00, 0x00, 0x10, 0x01, 0x00, 0x00, 0x40, 0x00, 0x30,
    0x47, 0x53, 0x55, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  ]);

const GUID = '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}';

function modelWith(embeddedFonts?: EmbeddedFontRef[]): DocxDocumentModel {
  const emptyHf = { default: null, first: null, even: null };
  return {
    section: {} as DocxDocumentModel['section'],
    body: [],
    headers: emptyHf,
    footers: emptyHf,
    embeddedFonts,
  };
}

describe('loadEmbeddedFonts (ECMA-376 §17.8.1 / §17.8.3)', () => {
  it('maps a 4-slot font to 4 faces with the correct weight/style descriptors', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Ubuntu', style: 'regular', partPath: 'word/fonts/font1.odttf', fontKey: GUID },
      { fontName: 'Ubuntu', style: 'bold', partPath: 'word/fonts/font2.odttf', fontKey: GUID },
      { fontName: 'Ubuntu', style: 'italic', partPath: 'word/fonts/font3.odttf', fontKey: GUID },
      { fontName: 'Ubuntu', style: 'boldItalic', partPath: 'word/fonts/font4.odttf', fontKey: GUID },
    ];
    // Every part is a valid header obfuscated with the GUID (so de-obfuscation
    // yields a valid sfnt), keyed by path.
    const bytesByPath = new Map(
      refs.map((r) => [r.partPath, deobfuscateOdttf(validHeader(), GUID)]),
    );
    await loadEmbeddedFonts(modelWith(refs), async (p) => bytesByPath.get(p)!);

    expect(added).toHaveLength(4);
    expect(added.every((f) => f.family === 'Ubuntu')).toBe(true);
    const byDesc = added.map((f) => `${f.descriptors?.weight}/${f.descriptors?.style}`).sort();
    expect(byDesc).toEqual([
      'bold/italic',
      'bold/normal',
      'normal/italic',
      'normal/normal',
    ]);
  });

  it('sets odttf=true for a .odttf part (bytes de-obfuscated to plaintext sfnt)', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Ubuntu', style: 'regular', partPath: 'word/fonts/font1.ODTTF', fontKey: GUID },
    ];
    // Obfuscated on the wire; loadEmbeddedFonts must flag odttf (case-insensitive
    // on the extension) so registerEmbeddedFonts de-obfuscates before FontFace.
    await loadEmbeddedFonts(modelWith(refs), async () => deobfuscateOdttf(validHeader(), GUID));
    expect(added).toHaveLength(1);
    // The first 4 bytes are the plaintext sfnt tag after de-obfuscation.
    expect(Array.from(new Uint8Array(added[0].source).slice(0, 4))).toEqual([
      0x00, 0x01, 0x00, 0x00,
    ]);
  });

  it('does not de-obfuscate a non-.odttf part (odttf=false)', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Roboto', style: 'regular', partPath: 'word/fonts/font1.ttf', fontKey: '' },
    ];
    // A raw sfnt part: odttf must be false so the bytes reach FontFace verbatim.
    await loadEmbeddedFonts(modelWith(refs), async () => validHeader());
    expect(added).toHaveLength(1);
    expect(Array.from(new Uint8Array(added[0].source).slice(0, 4))).toEqual([
      0x00, 0x01, 0x00, 0x00,
    ]);
  });

  it('skips a face whose fetch rejects, keeping the rest', async () => {
    const added = installFontFaceSet();
    const refs: EmbeddedFontRef[] = [
      { fontName: 'Good', style: 'regular', partPath: 'word/fonts/good.ttf', fontKey: '' },
      { fontName: 'Missing', style: 'regular', partPath: 'word/fonts/missing.ttf', fontKey: '' },
    ];
    await loadEmbeddedFonts(modelWith(refs), async (p) => {
      if (p.endsWith('missing.ttf')) throw new Error('no such part');
      return validHeader();
    });
    expect(added.map((f) => f.family)).toEqual(['Good']);
  });

  it('no-ops (no fetch) when embeddedFonts is empty or undefined', async () => {
    installFontFaceSet();
    const fetchSpy = vi.fn(async () => validHeader());

    await loadEmbeddedFonts(modelWith([]), fetchSpy);
    expect(fetchSpy).not.toHaveBeenCalled();

    await loadEmbeddedFonts(modelWith(undefined), fetchSpy);
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
