import { describe, it, expect, vi } from 'vitest';
import { PptxPresentation } from './presentation';
import type { WorkerRequest } from './types';

/**
 * `PptxPresentation.getImage(path, mime)` routes through the persistent worker
 * via the `extractImage` message (twin of `getMedia`/`extractMedia`), wraps the
 * returned bytes in a Blob of the requested MIME, and serves repeat calls from
 * its per-instance cache so the worker is hit at most once per path.
 *
 * The constructor opens a real Worker, so we build the instance off-prototype
 * and inject a fake `_bridge` whose `request` resolves an `imageExtracted`
 * response. This isolates the cache + Blob-wrapping contract from the worker.
 */
/** The subset of PptxPresentation this test exercises. Kept separate from the
 *  class type so the off-prototype build doesn't intersect its private fields
 *  (which would collapse to `never` under `tsc`). */
interface GetImageProbe {
  getImage(imagePath: string, mimeType: string): Promise<Blob>;
}

describe('PptxPresentation.getImage', () => {
  function makePresentation(requestImpl: (req: WorkerRequest) => unknown) {
    const request = vi.fn((build: (id: number) => WorkerRequest) =>
      Promise.resolve(requestImpl(build(1))),
    );
    // Build off the real prototype (so the real getImage runs) but inject only
    // the private collaborators it touches. Cast through unknown to avoid
    // intersecting the class's private members.
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._bridge = { request };
    instance._imageCache = new Map<string, Promise<Blob>>();
    const pres = instance as unknown as GetImageProbe;
    return { pres, request };
  }

  const bytesFor = (s: string) => new TextEncoder().encode(s).buffer;

  it('wraps extracted bytes in a Blob of the requested MIME', async () => {
    const payload = bytesFor('PNGDATA');
    const { pres, request } = makePresentation((req) => {
      expect(req.kind).toBe('extractImage');
      expect((req as Extract<WorkerRequest, { kind: 'extractImage' }>).path).toBe(
        'ppt/media/image1.png',
      );
      return { kind: 'imageExtracted', id: 1, bytes: payload };
    });

    const blob = await pres.getImage('ppt/media/image1.png', 'image/png');
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe('image/png');
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(
      new Uint8Array(bytesFor('PNGDATA')),
    );
    expect(request).toHaveBeenCalledTimes(1);
  });

  it('serves a second call for the same path from cache (one worker request)', async () => {
    const { pres, request } = makePresentation(() => ({
      kind: 'imageExtracted',
      id: 1,
      bytes: bytesFor('X'),
    }));

    const a = await pres.getImage('ppt/media/image1.png', 'image/png');
    const b = await pres.getImage('ppt/media/image1.png', 'image/png');
    // Same cached promise → identical Blob, single underlying request.
    expect(a).toBe(b);
    expect(request).toHaveBeenCalledTimes(1);
  });
});
