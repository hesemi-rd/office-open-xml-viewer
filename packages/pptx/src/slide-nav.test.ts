import { describe, it, expect } from 'vitest';
import {
  buildSlidePartIndex,
  resolveSlidePartTarget,
  resolveInternalSlideTarget,
} from './slide-nav';

// A three-slide deck's part names in sldIdLst order, as the parser stamps them.
const PART_NAMES = ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml'];

describe('buildSlidePartIndex', () => {
  it('maps each part name to its sldIdLst-order index', () => {
    const idx = buildSlidePartIndex(PART_NAMES);
    expect(idx.get('ppt/slides/slide1.xml')).toBe(0);
    expect(idx.get('ppt/slides/slide2.xml')).toBe(1);
    expect(idx.get('ppt/slides/slide3.xml')).toBe(2);
  });

  it('skips undefined / empty entries and keeps the first on a duplicate', () => {
    const idx = buildSlidePartIndex([
      'ppt/slides/slide1.xml',
      undefined,
      '',
      'ppt/slides/slide1.xml', // duplicate — first (index 0) wins
    ]);
    expect(idx.size).toBe(1);
    expect(idx.get('ppt/slides/slide1.xml')).toBe(0);
  });
});

describe('resolveSlidePartTarget', () => {
  const idx = buildSlidePartIndex(PART_NAMES);

  it('resolves a slide-rel target (../slides/slideN.xml) to its index', () => {
    // The exact hlinksldjump rel Target: authored relative to ppt/slides.
    expect(resolveSlidePartTarget('../slides/slide3.xml', idx)).toBe(2);
    expect(resolveSlidePartTarget('../slides/slide1.xml', idx)).toBe(0);
  });

  it('resolves a root-absolute slide target', () => {
    expect(resolveSlidePartTarget('/ppt/slides/slide2.xml', idx)).toBe(1);
  });

  it('returns undefined for a target that names no known slide', () => {
    expect(resolveSlidePartTarget('../slides/slide9.xml', idx)).toBeUndefined();
    expect(resolveSlidePartTarget('https://example.com/', idx)).toBeUndefined();
    expect(resolveSlidePartTarget('', idx)).toBeUndefined();
  });
});

describe('resolveInternalSlideTarget', () => {
  const idx = buildSlidePartIndex(PART_NAMES);

  it('resolves a specific slide-part jump', () => {
    expect(resolveInternalSlideTarget('../slides/slide3.xml', idx, 0)).toBe(2);
  });

  it('resolves relative show-jump verbs against the current slide', () => {
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=firstslide', idx, 1)).toBe(0);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=lastslide', idx, 1)).toBe(2);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=nextslide', idx, 1)).toBe(2);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=previousslide', idx, 1)).toBe(
      0,
    );
  });

  it('clamps nextslide/previousslide at the deck ends', () => {
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=nextslide', idx, 2)).toBe(2);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=previousslide', idx, 0)).toBe(
      0,
    );
  });

  it('returns undefined for an unresolvable internal ref', () => {
    // A show-action verb we don't navigate on, and no part match either.
    expect(
      resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=endshow', idx, 0),
    ).toBeUndefined();
  });
});
