import { describe, expect, it } from 'vitest';
import type { SourceRef } from './types.js';
import {
  createImageMetadataService,
  createMathMetadataService,
  imageResourceKey,
  mathResourceKey,
} from './resources.js';

const source: SourceRef = { story: 'body', storyInstance: 'body', path: [2, 1] };

describe('layout resource snapshots', () => {
  it('uses stable source-derived string keys and returns plain image metadata', () => {
    const key = imageResourceKey(source, 'word/media/image1.png');
    const service = createImageMetadataService([{
      resourceKey: key,
      widthPt: 72,
      heightPt: 36,
      mimeType: 'image/png',
    }]);

    expect(key).toBe('image:body:body:2.1:word%2Fmedia%2Fimage1.png');
    expect(service.resolve(key)).toEqual({ widthPt: 72, heightPt: 36, mimeType: 'image/png' });
    expect(Object.getPrototypeOf(service.resolve(key))).toBe(Object.prototype);
    expect(() => service.resolve('image:missing')).toThrow(/Unknown image resource/);
  });

  it('keys math by source rather than parser object identity', () => {
    const first = mathResourceKey(source, 'formula');
    const equivalentSource: SourceRef = { story: 'body', storyInstance: 'body', path: [2, 1] };
    const second = mathResourceKey(equivalentSource, 'formula');
    const service = createMathMetadataService([{
      resourceKey: first,
      widthEm: 2.5,
      ascentEm: 0.8,
      descentEm: 0.2,
      diagnostics: [],
    }]);

    expect(first).toBe(second);
    expect(service.resolve(second)).toEqual({
      resourceKey: first,
      widthEm: 2.5,
      ascentEm: 0.8,
      descentEm: 0.2,
      diagnostics: [],
    });
    expect(service.fingerprint).toMatch(/^math:/);
  });

  it('fingerprints immutable resource content independent of insertion order', () => {
    const a = { resourceKey: 'a', widthPt: 1, heightPt: 2, mimeType: 'image/png' };
    const b = { resourceKey: 'b', widthPt: 3, heightPt: 4, mimeType: 'image/jpeg' };
    const forward = createImageMetadataService([a, b]);
    const reverse = createImageMetadataService([b, a]);

    expect(forward.fingerprint).toBe(reverse.fingerprint);
    expect(Object.isFrozen(forward.resolve('a'))).toBe(true);
  });
});
