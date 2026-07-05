import { describe, it, expect } from 'vitest';
import {
  resolveOpcPartName,
  parseRelativeSlideJump,
  resolveRelativeSlideJump,
} from './internal-target';

describe('resolveOpcPartName', () => {
  it('resolves a slide-rel relative target against ppt/slides (the pptx jump case)', () => {
    // The exact string a pptx internal slide jump carries: <a:hlinkClick r:id>
    // -> rels Target "../slides/slide3.xml", authored relative to ppt/slides.
    expect(resolveOpcPartName('ppt/slides', '../slides/slide3.xml')).toBe('ppt/slides/slide3.xml');
  });

  it('resolves a bare sibling target', () => {
    expect(resolveOpcPartName('ppt/slides', 'slide1.xml')).toBe('ppt/slides/slide1.xml');
  });

  it('root-absolute target ignores baseDir and drops the leading slash', () => {
    expect(resolveOpcPartName('ppt/slides', '/ppt/slides/slide2.xml')).toBe(
      'ppt/slides/slide2.xml',
    );
  });

  it('normalizes multi-level ..', () => {
    expect(resolveOpcPartName('ppt/slides/sub', '../../slides/slide4.xml')).toBe(
      'ppt/slides/slide4.xml',
    );
  });

  it('matches the Rust resolve_target normalization for a media-style target', () => {
    expect(resolveOpcPartName('ppt/slides', '../media/image1.png')).toBe('ppt/media/image1.png');
  });
});

describe('parseRelativeSlideJump', () => {
  it('parses each of the four navigation verbs', () => {
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=firstslide')).toBe('firstslide');
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=lastslide')).toBe('lastslide');
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=nextslide')).toBe('nextslide');
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=previousslide')).toBe(
      'previousslide',
    );
  });

  it('is case-insensitive on the verb', () => {
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=FirstSlide')).toBe('firstslide');
  });

  it('returns null for a non-navigation show action', () => {
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=endshow')).toBeNull();
    expect(parseRelativeSlideJump('ppaction://hlinkshowjump?jump=lastslideviewed')).toBeNull();
  });

  it('returns null for a non-show action (e.g. a slide-part jump verb)', () => {
    expect(parseRelativeSlideJump('ppaction://hlinksldjump')).toBeNull();
  });
});

describe('resolveRelativeSlideJump', () => {
  it('firstslide / lastslide land on the deck ends regardless of current', () => {
    expect(resolveRelativeSlideJump('firstslide', 3, 10)).toBe(0);
    expect(resolveRelativeSlideJump('lastslide', 3, 10)).toBe(9);
  });

  it('nextslide / previousslide step by one', () => {
    expect(resolveRelativeSlideJump('nextslide', 3, 10)).toBe(4);
    expect(resolveRelativeSlideJump('previousslide', 3, 10)).toBe(2);
  });

  it('clamps at the boundaries rather than wrapping', () => {
    expect(resolveRelativeSlideJump('nextslide', 9, 10)).toBe(9);
    expect(resolveRelativeSlideJump('previousslide', 0, 10)).toBe(0);
  });

  it('returns undefined for an empty deck', () => {
    expect(resolveRelativeSlideJump('firstslide', 0, 0)).toBeUndefined();
  });
});
