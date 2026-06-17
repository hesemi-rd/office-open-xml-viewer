import { describe, it, expect, afterEach, vi } from 'vitest';
import { getCachedSvgImageByPath, dropSvgImageCache } from './svg-image-by-path';

describe('getCachedSvgImageByPath', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('fetches bytes, makes an object URL, loads an <img>, dedupes by path', async () => {
    let created = 0;
    vi.stubGlobal('URL', { createObjectURL: () => { created++; return `blob:${created}`; },
                           revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    const b = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    expect(a).toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(created).toBe(1);
  });

  it('self-evicts on failure (no poisoned cache) and revokes the failed object URL', async () => {
    let created = 0;
    let revoked = 0;
    vi.stubGlobal('URL', {
      createObjectURL: () => { created++; return `blob:${created}`; },
      revokeObjectURL: () => { revoked++; },
    });
    // <img> that always fails to load.
    class FailImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onerror && this.onerror()); } }
    vi.stubGlobal('Image', FailImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    await expect(getCachedSvgImageByPath('p.svg', fetchImage)).rejects.toThrow();
    expect(revoked).toBe(1); // failed URL revoked, not leaked
    // Second call must RETRY (cache self-evicted), not return a cached rejection.
    await expect(getCachedSvgImageByPath('p.svg', fetchImage)).rejects.toThrow();
    expect(fetchImage).toHaveBeenCalledTimes(2);
  });

  it('namespaces the cache by fetchImage — same zip path from two documents does not cross-contaminate', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    // Two different documents reference the SAME internal zip path. Their byte
    // sources (fetchImage closures) differ — the cache must not serve doc A's
    // decoded SVG for doc B's request.
    const fetchA = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const fetchB = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/image1.svg', fetchA);
    const b = await getCachedSvgImageByPath('word/media/image1.svg', fetchB);
    expect(fetchA).toHaveBeenCalledTimes(1);
    expect(fetchB).toHaveBeenCalledTimes(1); // B must consult its OWN source, not hit A's cache
    expect(a).not.toBe(b); // distinct decoded images
    // Within one document, the path still dedupes.
    const a2 = await getCachedSvgImageByPath('word/media/image1.svg', fetchA);
    expect(a2).toBe(a);
    expect(fetchA).toHaveBeenCalledTimes(1);
  });

  it('dropSvgImageCache revokes every object URL for a fetchImage and lets it re-decode', async () => {
    let revoked = 0;
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => { revoked++; } });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    await getCachedSvgImageByPath('a.svg', fetchImage);
    await getCachedSvgImageByPath('b.svg', fetchImage);
    dropSvgImageCache(fetchImage); // e.g. on Document.destroy()
    expect(revoked).toBe(2); // both live object URLs released
    await getCachedSvgImageByPath('a.svg', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3); // cache cleared → fresh decode
  });

  it('awaits img.decode() before resolving when available', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    let decoded = false;
    class DecodeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      decode() { return Promise.resolve().then(() => { decoded = true; }); }
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', DecodeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    await getCachedSvgImageByPath('d.svg', fetchImage);
    expect(decoded).toBe(true); // resolved only after decode() completed
  });
});
