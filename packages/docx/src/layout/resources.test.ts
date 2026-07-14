import { describe, expect, it } from 'vitest';
import type { SourceRef } from './types.js';
import {
  createImageMetadataService,
  createMathMetadataService,
  documentImageMetadataRecords,
  imageResourceKey,
  mathResourceKey,
} from './resources.js';
import { stableFingerprint } from './fingerprint.js';
import type { DocxDocumentModel } from '../types.js';

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

  it('keeps repeated inline/display formulas as distinct source occurrences', () => {
    const inlineA = mathResourceKey({ story: 'body', storyInstance: 'body', path: [0, 0] }, 'inline');
    const inlineB = mathResourceKey({ story: 'body', storyInstance: 'body', path: [1, 0] }, 'inline');
    const display = mathResourceKey({ story: 'body', storyInstance: 'body', path: [2, 0] }, 'display');
    const service = createMathMetadataService([
      { resourceKey: inlineA, widthEm: 1, ascentEm: 1, descentEm: 0, diagnostics: [] },
      { resourceKey: inlineB, widthEm: 1, ascentEm: 1, descentEm: 0, diagnostics: [] },
      { resourceKey: display, widthEm: 2, ascentEm: 1, descentEm: 0, diagnostics: [] },
    ]);

    expect(new Set([inlineA, inlineB, display]).size).toBe(3);
    expect(service.resolve(inlineA).resourceKey).toBe(inlineA);
    expect(service.resolve(inlineB).resourceKey).toBe(inlineB);
    expect(() => service.resolve('formula:inline')).toThrow(/Unknown math resource/);
  });

  it('fingerprints immutable resource content independent of insertion order', () => {
    const a = { resourceKey: 'a', widthPt: 1, heightPt: 2, mimeType: 'image/png' };
    const b = { resourceKey: 'b', widthPt: 3, heightPt: 4, mimeType: 'image/jpeg' };
    const forward = createImageMetadataService([a, b]);
    const reverse = createImageMetadataService([b, a]);

    expect(forward.fingerprint).toBe(reverse.fingerprint);
    expect(Object.isFrozen(forward.resolve('a'))).toBe(true);
  });

  it('does not use the legacy 32-bit digest as cache identity', () => {
    expect(stableFingerprint('collision', 'j4plhpxqpj'))
      .not.toBe(stableFingerprint('collision', '9rndrgx1elo'));
  });

  it('deep-freezes math metadata and diagnostics and rejects duplicate identities', () => {
    const diagnostic = { code: 'UNSUPPORTED_FEATURE' as const, severity: 'warning' as const, message: 'fallback' };
    const service = createMathMetadataService([{
      resourceKey: 'math:a', widthEm: 1, ascentEm: 0.8, descentEm: 0.2, diagnostics: [diagnostic],
    }]);
    const resolved = service.resolve('math:a');

    expect(Object.isFrozen(resolved)).toBe(true);
    expect(Object.isFrozen(resolved.diagnostics)).toBe(true);
    expect(Object.isFrozen(resolved.diagnostics[0])).toBe(true);
    expect(() => createMathMetadataService([
      { resourceKey: 'math:a', widthEm: 1, ascentEm: 1, descentEm: 0, diagnostics: [] },
      { resourceKey: 'math:a', widthEm: 2, ascentEm: 1, descentEm: 0, diagnostics: [] },
    ])).toThrow(/Duplicate math resource key/);
  });

  it('enumerates repeated image parts with typed story and textbox occurrence keys', () => {
    const image = (widthPt: number, heightPt: number) => ({
      type: 'image', imagePath: 'word/media/shared.png', mimeType: 'image/png', widthPt, heightPt,
    });
    const paragraph = (runs: unknown[], numbering: unknown = null) => ({ type: 'paragraph', runs, numbering });
    const doc = {
      body: [paragraph([
        image(10, 20),
        {
          type: 'shape', widthPt: 30, heightPt: 40, textBlocks: [{
            text: '', imagePath: 'word/media/shared.png', mimeType: 'image/png', imageWidthPt: 5, imageHeightPt: 6,
          }],
        },
      ], {
        picBulletImagePath: 'word/media/bullet.gif', picBulletMimeType: 'image/gif',
        picBulletWidthPt: 7, picBulletHeightPt: 8,
      })],
      headers: { default: { body: [paragraph([image(11, 21)])] }, first: null, even: null },
      footers: { default: { body: [paragraph([image(12, 22)])] }, first: null, even: null },
      footnotes: [{ id: '4', content: [paragraph([image(13, 23)])] }],
      endnotes: [{ id: '9', content: [paragraph([image(14, 24)])] }],
    } as unknown as DocxDocumentModel;

    const records = documentImageMetadataRecords(doc);
    expect(records).toEqual(expect.arrayContaining([
      { resourceKey: imageResourceKey({ story: 'body', storyInstance: 'body', path: [0, 0] }, 'word/media/shared.png'), widthPt: 10, heightPt: 20, mimeType: 'image/png' },
      { resourceKey: imageResourceKey({ story: 'header', storyInstance: 'default', path: [0, 0] }, 'word/media/shared.png'), widthPt: 11, heightPt: 21, mimeType: 'image/png' },
      { resourceKey: imageResourceKey({ story: 'footer', storyInstance: 'default', path: [0, 0] }, 'word/media/shared.png'), widthPt: 12, heightPt: 22, mimeType: 'image/png' },
      { resourceKey: imageResourceKey({ story: 'textbox', storyInstance: 'body:body:0.1', path: [0] }, 'word/media/shared.png'), widthPt: 5, heightPt: 6, mimeType: 'image/png' },
      { resourceKey: imageResourceKey({ story: 'body', storyInstance: 'body', path: [0] }, 'word/media/bullet.gif'), widthPt: 7, heightPt: 8, mimeType: 'image/gif' },
      { resourceKey: imageResourceKey({ story: 'footnote', storyInstance: '4', path: [0, 0] }, 'word/media/shared.png'), widthPt: 13, heightPt: 23, mimeType: 'image/png' },
      { resourceKey: imageResourceKey({ story: 'endnote', storyInstance: '9', path: [0, 0] }, 'word/media/shared.png'), widthPt: 14, heightPt: 24, mimeType: 'image/png' },
    ]));
    expect(new Set(records.map((record) => record.resourceKey)).size).toBe(records.length);
  });
});
