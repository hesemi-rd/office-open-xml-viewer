import { describe, it, expect } from 'vitest';
import { classifyPptxHyperlink } from './hyperlink.js';

// IX1 — pptx hyperlink classification. The Rust parser hands the TS side a
// single resolved target string per run/shape plus (for shapes) the raw
// `<a:hlinkClick @action>`. classifyPptxHyperlink turns that pair into the
// shared HyperlinkTarget shape used by the overlay + viewer. ECMA-376
// §21.1.2.3.5: an action verb (ppaction://…) or a scheme-less/internal part
// name is INTERNAL; a navigable web scheme is EXTERNAL.
describe('classifyPptxHyperlink', () => {
  it('returns undefined when there is no link', () => {
    expect(classifyPptxHyperlink(undefined)).toBeUndefined();
    expect(classifyPptxHyperlink('')).toBeUndefined();
    expect(classifyPptxHyperlink('', '')).toBeUndefined();
  });

  it('classifies a navigable web URL as external', () => {
    expect(classifyPptxHyperlink('https://example.com/a')).toEqual({ kind: 'external', url: 'https://example.com/a' });
    expect(classifyPptxHyperlink('http://x')).toEqual({ kind: 'external', url: 'http://x' });
    expect(classifyPptxHyperlink('mailto:a@b.com')).toEqual({ kind: 'external', url: 'mailto:a@b.com' });
    expect(classifyPptxHyperlink('tel:+15551234567')).toEqual({ kind: 'external', url: 'tel:+15551234567' });
  });

  it('classifies a ppaction:// slide-jump as internal, preferring the resolved part name', () => {
    // Resolved rels target is the slide part; action marks it internal.
    expect(
      classifyPptxHyperlink('../slides/slide3.xml', 'ppaction://hlinksldjump'),
    ).toEqual({ kind: 'internal', ref: '../slides/slide3.xml' });
    // Action present but no resolved target (rels miss): fall back to the verb.
    expect(
      classifyPptxHyperlink(undefined, 'ppaction://hlinkshowjump?jump=firstslide'),
    ).toEqual({ kind: 'internal', ref: 'ppaction://hlinkshowjump?jump=firstslide' });
  });

  it('classifies a bare internal part name (no scheme, no action) as internal', () => {
    expect(classifyPptxHyperlink('../slides/slide2.xml')).toEqual({ kind: 'internal', ref: '../slides/slide2.xml' });
  });

  it('treats a non-navigable scheme (would be blocked externally) as internal, matching the sanitiser boundary', () => {
    // The viewer's default open would refuse these; classifying them internal
    // (rather than external) means the overlay never hands a blocked scheme to
    // openExternalHyperlink. The parser is unlikely to resolve such a target,
    // but the boundary must line up with DEFAULT_ALLOWED_HYPERLINK_SCHEMES.
    expect(classifyPptxHyperlink('javascript:alert(1)')).toEqual({ kind: 'internal', ref: 'javascript:alert(1)' });
    expect(classifyPptxHyperlink('file:///etc/passwd')).toEqual({ kind: 'internal', ref: 'file:///etc/passwd' });
  });
});
