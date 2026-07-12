import { describe, expect, it } from 'vitest';
import {
  __test_physicalLayoutSection as physicalLayoutSection,
  __test_verticalLayoutSection as verticalLayoutSection,
} from './renderer.js';
import type { SectionProps } from './types.js';

// ECMA-376 §17.6.20 + §17.10.1 — a vertical (tbRl) section's header/footer stay
// HORIZONTAL at the physical top/bottom margins (Word ground truth, issue #988's
// batch-3 adjudication). The body is laid out in the SWAPPED logical geometry and
// the page paint is rotated +90°; the header/footer must be drawn back in the
// physical page frame, which `physicalLayoutSection` (the inverse of
// `verticalLayoutSection`) recovers.

function makeSection(over: Partial<SectionProps>): SectionProps {
  return {
    pageWidth: 612,
    pageHeight: 792,
    marginTop: 72,
    marginRight: 90,
    marginBottom: 108,
    marginLeft: 126,
    headerDistance: 36,
    footerDistance: 36,
    titlePage: false,
    evenAndOddHeaders: false,
    textDirection: 'tbRl',
    ...over,
  };
}

describe('vertical header/footer physical frame (§17.6.20)', () => {
  it('physicalLayoutSection inverts verticalLayoutSection (round-trip)', () => {
    const phys = makeSection({});
    const round = physicalLayoutSection(verticalLayoutSection(phys));
    expect(round.pageWidth).toBe(phys.pageWidth);
    expect(round.pageHeight).toBe(phys.pageHeight);
    expect(round.marginTop).toBe(phys.marginTop);
    expect(round.marginRight).toBe(phys.marginRight);
    expect(round.marginBottom).toBe(phys.marginBottom);
    expect(round.marginLeft).toBe(phys.marginLeft);
    expect(round.headerDistance).toBe(phys.headerDistance);
    expect(round.footerDistance).toBe(phys.footerDistance);
  });

  it('recovers the physical page box + margins from the logical (swapped) section', () => {
    // The renderer lays a vertical section out in `verticalLayoutSection(phys)`;
    // the header/footer path takes THAT logical section and must recover the
    // physical page box (612×792) with the four margins on their physical edges.
    const phys = makeSection({});
    const logical = verticalLayoutSection(phys);
    // Sanity: the logical section is swapped (width = physical height).
    expect(logical.pageWidth).toBe(792);
    expect(logical.pageHeight).toBe(612);

    const recovered = physicalLayoutSection(logical);
    // Physical page box: portrait letter.
    expect(recovered.pageWidth).toBe(612);
    expect(recovered.pageHeight).toBe(792);
    // Physical margins land on their original edges.
    expect(recovered.marginTop).toBe(72);
    expect(recovered.marginRight).toBe(90);
    expect(recovered.marginBottom).toBe(108);
    expect(recovered.marginLeft).toBe(126);
  });

  it('preserves non-geometry fields (textDirection, docGrid) so the frame stays vertical-aware', () => {
    const phys = makeSection({ docGridType: 'lines', docGridLinePitch: 18 });
    const recovered = physicalLayoutSection(verticalLayoutSection(phys));
    expect(recovered.textDirection).toBe('tbRl');
    expect(recovered.docGridType).toBe('lines');
    expect(recovered.docGridLinePitch).toBe(18);
  });
});
