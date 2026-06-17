import { describe, it, expect, vi } from 'vitest';
import { DocxDocument } from './document';
import type { WorkerRequest } from './types';

/**
 * `DocxDocument.getImage(path, mime)` routes through the persistent worker via
 * the `extractImage` message (docx uses the `type` discriminant), wraps the
 * returned bytes in a Blob of the requested MIME, and serves repeat calls from
 * its per-instance cache so the worker is hit at most once per path. Mirrors
 * pptx's `presentation.image.test.ts`.
 *
 * The constructor opens a real Worker, so we build the instance off-prototype
 * and inject a fake `_bridge` whose `request` resolves an `imageExtracted`
 * response. This isolates the cache + Blob-wrapping contract from the worker.
 */
/** The subset of DocxDocument this test exercises. Kept separate from the class
 *  type so the off-prototype build doesn't intersect its private fields (which
 *  would collapse to `never` under `tsc`). */
interface GetImageProbe {
  getImage(imagePath: string, mimeType: string): Promise<Blob>;
}

describe('DocxDocument.getImage', () => {
  function makeDocument(requestImpl: (req: WorkerRequest) => unknown) {
    const request = vi.fn((build: (id: number) => WorkerRequest) =>
      Promise.resolve(requestImpl(build(1))),
    );
    // Build off the real prototype (so the real getImage runs) but inject only
    // the private collaborators it touches. Cast through unknown to avoid
    // intersecting the class's private members.
    const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
    instance._bridge = { request };
    instance._imageCache = new Map<string, Promise<Blob>>();
    const doc = instance as unknown as GetImageProbe;
    return { doc, request };
  }

  const bytesFor = (s: string) => new TextEncoder().encode(s).buffer;

  it('wraps extracted bytes in a Blob of the requested MIME', async () => {
    const payload = bytesFor('PNGDATA');
    const { doc, request } = makeDocument((req) => {
      expect(req.type).toBe('extractImage');
      expect((req as Extract<WorkerRequest, { type: 'extractImage' }>).path).toBe(
        'word/media/image1.png',
      );
      return { type: 'imageExtracted', id: 1, bytes: payload };
    });

    const blob = await doc.getImage('word/media/image1.png', 'image/png');
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe('image/png');
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(
      new Uint8Array(bytesFor('PNGDATA')),
    );
    expect(request).toHaveBeenCalledTimes(1);
  });

  it('serves a second call for the same path from cache (one worker request)', async () => {
    const { doc, request } = makeDocument(() => ({
      type: 'imageExtracted',
      id: 1,
      bytes: bytesFor('X'),
    }));

    const a = await doc.getImage('word/media/image1.png', 'image/png');
    const b = await doc.getImage('word/media/image1.png', 'image/png');
    // Same cached promise → identical Blob, single underlying request.
    expect(a).toBe(b);
    expect(request).toHaveBeenCalledTimes(1);
  });
});
