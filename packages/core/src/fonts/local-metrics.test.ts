import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  loadLocalFontMetrics,
  unloadLocalFontMetrics,
} from './local-metrics.js';
import { _resetFontRegistryForTests } from './font-registry.js';

const G = globalThis as unknown as Record<string, unknown>;
const ORIGINALS = {
  document: G.document,
  self: G.self,
  FontFace: G.FontFace,
  OffscreenCanvas: G.OffscreenCanvas,
};

beforeEach(() => {
  _resetFontRegistryForTests();
});

afterEach(() => {
  for (const [key, value] of Object.entries(ORIGINALS)) {
    if (value === undefined) delete G[key];
    else G[key] = value;
  }
  _resetFontRegistryForTests();
});

function installFontEnvironment(widthOf: (text: string) => number = () => 50) {
  const added: FakeFontFace[] = [];
  const deleted: FakeFontFace[] = [];
  class FakeFontFace {
    status: FontFaceLoadStatus = 'unloaded';
    constructor(
      readonly family: string,
      readonly source: string,
      readonly descriptors: FontFaceDescriptors = {},
    ) {}
    async load(): Promise<this> {
      if (this.source.includes('Missing Face')) {
        this.status = 'error';
        throw new DOMException('missing', 'NetworkError');
      }
      this.status = 'loaded';
      return this;
    }
  }
  const fonts = {
    add(face: FakeFontFace) { added.push(face); },
    delete(face: FakeFontFace) { deleted.push(face); return true; },
  } as unknown as FontFaceSet;
  class FakeOffscreenCanvas {
    constructor(_width: number, _height: number) {}
    getContext() {
      return {
        font: '',
        measureText: (text: string) => ({
          width: widthOf(text),
          fontBoundingBoxAscent: 106,
          fontBoundingBoxDescent: 44,
        }),
      };
    }
  }
  G.document = { fonts };
  G.FontFace = FakeFontFace;
  G.OffscreenCanvas = FakeOffscreenCanvas;
  return { added, deleted };
}

describe('loadLocalFontMetrics', () => {
  it('loads an exact local face under an isolated alias and derives its line ratio', async () => {
    const { added, deleted } = installFontEnvironment();

    const loaded = await loadLocalFontMetrics([{
      family: 'メイリオ',
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    }]);

    expect(loaded.faces).toHaveLength(1);
    expect(added).toHaveLength(1);
    expect(added[0].source).toBe('local("Meiryo")');
    expect(loaded.metrics['メイリオ'].family).toBe(added[0].family);
    expect(loaded.metrics['メイリオ'].lineHeightRatio).toBeCloseTo(1.95, 8);

    unloadLocalFontMetrics(loaded.faces);
    expect(deleted).toEqual([added[0]]);
  });

  it('does not mistake a missing local font for a fallback face', async () => {
    const { added, deleted } = installFontEnvironment();

    const loaded = await loadLocalFontMetrics([{
      family: 'Missing',
      localNames: ['Missing Face'],
      lineHeightMultiplier: 1.3,
    }]);

    expect(loaded.faces).toEqual([]);
    expect(loaded.metrics).toEqual({});
    expect(added).toHaveLength(1);
    expect(deleted).toEqual(added);
  });

  it('maps styled Canvas tuples to one positively loaded regular alias and fingerprints synthesis', async () => {
    const { added } = installFontEnvironment();
    const loaded = await loadLocalFontMetrics([
      { family: 'Authored Serif', localNames: ['Authored Serif'] },
      { family: 'Authored Serif', localNames: ['Authored Serif'], weight: 700, style: 'italic' },
    ]);

    expect(added).toHaveLength(1);
    expect(added[0].descriptors).toEqual({});
    expect(loaded.metrics['authored serif'].family)
      .toBe(loaded.metrics['authored serif:700:italic'].family);
    expect(loaded.metrics['authored serif:700:italic']).toMatchObject({
      requestedFamily: 'Authored Serif', weight: 700, style: 'italic',
    });
    expect(loaded.metrics['authored serif:700:italic'].geometrySignature).toMatch(/^\[/);
    expect(loaded.metrics['authored serif:700:italic'].lineHeightRatio).toBeUndefined();
  });

  it('changes the geometry hash when document-used glyphs differ despite matching sentinels', async () => {
    const request = [{
      family: 'Authored Serif', localNames: ['Authored Serif'], geometryProbeTexts: ['Q'],
    }];
    installFontEnvironment((text) => text === 'Q' ? 51 : 50);
    const first = await loadLocalFontMetrics(request);
    unloadLocalFontMetrics(first.faces);

    installFontEnvironment((text) => text === 'Q' ? 52 : 50);
    const second = await loadLocalFontMetrics(request);

    expect(first.metrics['authored serif'].geometrySignature)
      .not.toBe(second.metrics['authored serif'].geometrySignature);
  });

  it('waits for a shared face that another document is still loading', async () => {
    const added: Array<{ family: string; source: string; status: FontFaceLoadStatus }> = [];
    let finishLoad!: () => void;
    const loaded = new Promise<void>((resolve) => { finishLoad = resolve; });
    class DeferredFontFace {
      status: FontFaceLoadStatus = 'unloaded';
      constructor(readonly family: string, readonly source: string) {}
      async load(): Promise<this> {
        if (this.status === 'loaded') return this;
        this.status = 'loading';
        await loaded;
        this.status = 'loaded';
        return this;
      }
    }
    const fonts = {
      add(face: DeferredFontFace) { added.push(face); },
      delete() { return true; },
    } as unknown as FontFaceSet;
    class DeferredCanvas {
      getContext() {
        return {
          font: '',
          measureText: () => added[0]?.status === 'loaded'
            ? { width: 50, fontBoundingBoxAscent: 106, fontBoundingBoxDescent: 44 }
            : { width: Number.NaN, fontBoundingBoxAscent: Number.NaN, fontBoundingBoxDescent: Number.NaN },
        };
      }
    }
    G.document = { fonts };
    G.FontFace = DeferredFontFace;
    G.OffscreenCanvas = DeferredCanvas;
    const request = [{
      family: 'Meiryo',
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    }];

    const first = loadLocalFontMetrics(request);
    await Promise.resolve();
    const second = loadLocalFontMetrics(request);
    await Promise.resolve();
    finishLoad();

    const [a, b] = await Promise.all([first, second]);
    expect(added).toHaveLength(1);
    expect(a.metrics.meiryo.lineHeightRatio).toBeCloseTo(1.95, 8);
    expect(b.metrics.meiryo.lineHeightRatio).toBeCloseTo(1.95, 8);
  });

  it('is a no-op when the current context has no FontFaceSet', async () => {
    delete G.document;
    delete G.self;
    const loaded = await loadLocalFontMetrics([{
      family: 'Meiryo',
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    }]);
    expect(loaded).toEqual({ faces: [], metrics: {} });
  });
});
