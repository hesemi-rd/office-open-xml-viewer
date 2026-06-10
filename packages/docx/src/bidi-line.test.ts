import { describe, it, expect } from 'vitest';
import { segmentsHaveRtl, computeLineVisualOrder, resolveAlignEdge } from './bidi-line.js';

describe('segmentsHaveRtl', () => {
  it('detects Arabic/Hebrew, ignores Latin and objects', () => {
    expect(segmentsHaveRtl([{ text: 'Hello world' }])).toBe(false);
    expect(segmentsHaveRtl([{ text: 'price 99' }, {}])).toBe(false); // object seg has no text
    expect(segmentsHaveRtl([{ text: 'مرحبا' }])).toBe(true);
    expect(segmentsHaveRtl([{ text: 'abc ' }, { text: 'שלום' }])).toBe(true);
  });

  it('detects a run-level rtl mark even on neutral-only text', () => {
    // "1. " has no strong-RTL char, but the run carries <w:rtl> (§17.3.2.30).
    expect(segmentsHaveRtl([{ text: '1. ', rtl: true }])).toBe(true);
    expect(segmentsHaveRtl([{ text: '1. ' }])).toBe(false);
  });
});

  it('keeps LTR word order for English text in rtl-marked runs', () => {
    // w:rtl on a run of pure Latin (e.g. "first leader name" in sample-7's
    // signature block): the RLE embedding must NOT reverse the words — only
    // weak/neutral content (digits, punctuation) takes the RTL embedding.
    const segs = [
      { text: 'first ', rtl: true },
      { text: 'leader ', rtl: true },
      { text: 'name', rtl: true },
    ];
    const { order, rtl } = computeLineVisualOrder(segs, false);
    expect(order).toEqual([0, 1, 2]);
    expect(rtl).toEqual([false, false, false]);
  });

describe('resolveAlignEdge', () => {
  it('resolves logical start/end against base direction', () => {
    expect(resolveAlignEdge(undefined, false)).toBe('left');
    expect(resolveAlignEdge(undefined, true)).toBe('right');
    expect(resolveAlignEdge('start', true)).toBe('right');
    expect(resolveAlignEdge('end', true)).toBe('left');
    expect(resolveAlignEdge('end', false)).toBe('right');
    expect(resolveAlignEdge('center', true)).toBe('center');
    expect(resolveAlignEdge('both', true)).toBe('justify');
    // Transitional left/right are "semantically equivalent to start/end"
    // (ECMA-376 Part 4 §14.11.2) — logical, so they flip under an RTL base.
    expect(resolveAlignEdge('left', true)).toBe('right');
    expect(resolveAlignEdge('right', true)).toBe('left');
    expect(resolveAlignEdge('left', false)).toBe('left');
    expect(resolveAlignEdge('right', false)).toBe('right');
  });
});

describe('computeLineVisualOrder', () => {
  it('keeps pure-LTR segments in logical order', () => {
    const { order, rtl } = computeLineVisualOrder([{ text: 'Hello ' }, { text: 'world' }], false);
    expect(order).toEqual([0, 1]);
    expect(rtl).toEqual([false, false]);
  });

  it('reverses pure-RTL segments and marks them RTL', () => {
    const { order, rtl } = computeLineVisualOrder([{ text: 'שלום ' }, { text: 'עולם' }], true);
    expect(order).toEqual([1, 0]); // last word is visually leftmost
    expect(rtl).toEqual([true, true]);
  });

  it('orders an embedded LTR run inside an RTL line', () => {
    // logical: Hebrew, Latin, Hebrew (base RTL) -> visual L→R: [2,1,0]
    const { order, rtl } = computeLineVisualOrder(
      [{ text: 'שלום ' }, { text: 'world ' }, { text: 'טוב' }],
      true,
    );
    expect(order).toEqual([2, 1, 0]);
    expect(rtl[1]).toBe(false); // the Latin segment is LTR
    expect(rtl[0]).toBe(true);
    expect(rtl[2]).toBe(true);
  });

  it('resolves a leading-digit run-level rtl run with digits at the trailing (right) end', () => {
    // "1. " + Hebrew, both run-level rtl, base LTR (no w:bidi). Per §17.3.2.30
    // the digits must embed RTL, so visually the Hebrew is leftmost and the
    // "1. " (read RTL → ".1") sits at the right end.
    const { order, rtl } = computeLineVisualOrder(
      [{ text: '1. ', rtl: true }, { text: 'תוכן', rtl: true }],
      false,
    );
    expect(order).toEqual([1, 0]); // Hebrew visually left, "1." visually right
    expect(rtl).toEqual([true, true]);
  });

  it('keeps a trailing non-rtl numeric run leftmost in a base-RTL title (§17.3.2.30)', () => {
    // sample-7 title "المشاريع لعام 2026": three runs, the Arabic ones carry
    // w:rtl, "2026" does not, in a base-RTL (w:bidi) paragraph. Word renders
    // "2026" at the LEFT (logically last in the RTL line). A w:rtl run that
    // already has strong-RTL content must NOT be over-embedded above the base.
    const { order, rtl } = computeLineVisualOrder(
      [
        { text: 'المشاريع ', rtl: true },
        { text: 'لعام ', rtl: true },
        { text: '2026', rtl: false },
      ],
      true,
    );
    // `order` is logical indices in visual LEFT→RIGHT order, so the LEFTMOST
    // segment is order[0] (cf. the pure-RTL case where order=[1,0] puts the
    // logically-last word leftmost).
    expect(order).toEqual([2, 1, 0]);
    expect(order[0]).toBe(2); // "2026" visually leftmost
    expect(rtl[2]).toBe(false);
  });

  it('keeps the LTR fast path byte-identical when no rtl marks and no strong-RTL', () => {
    const { order, rtl } = computeLineVisualOrder(
      [{ text: '1. ' }, { text: 'item' }],
      false,
    );
    expect(order).toEqual([0, 1]);
    expect(rtl).toEqual([false, false]);
  });

  it('places a neutral object by surrounding direction', () => {
    // object between two Hebrew words, base RTL -> object stays in RTL flow
    const { order } = computeLineVisualOrder([{ text: 'שלום ' }, {}, { text: 'טוב' }], true);
    expect(order).toEqual([2, 1, 0]);
  });

  it('orders an AN-classified date to Word ordering (2026-02-28)', () => {
    // sample-7 date cell: logical "28-02-2026" in an Arabic complex-script run
    // (w:rtl, w:lang w:bidi="ae-AR"). The renderer pre-splits it into digit-
    // group / separator segments and tags each with `digitsAsAN`. Under the
    // RTL cell base, classifying the European digits as AN reorders the groups
    // to Word's "2026-02-28" (UAX#9 §4.3 HL1 higher-level protocol).
    const segs = [
      { text: '28', rtl: true, digitsAsAN: true },
      { text: '-', rtl: true, digitsAsAN: true },
      { text: '02', rtl: true, digitsAsAN: true },
      { text: '-', rtl: true, digitsAsAN: true },
      { text: '2026', rtl: true, digitsAsAN: true },
    ];
    const { order } = computeLineVisualOrder(segs, true);
    // Visual L→R: 2026 - 02 - 28  →  logical indices [4,3,2,1,0].
    const visual = order.map((i) => segs[i].text).join('');
    expect(visual).toBe('2026-02-28');
  });

  it('leaves a date as logical order when digits are NOT AN-classified (pure UAX#9 EN)', () => {
    // Same split, base RTL, but without the digitsAsAN flag: European digits
    // stay EN, all at the even embedded level, so the groups keep logical order.
    const segs = [
      { text: '28', rtl: true },
      { text: '-', rtl: true },
      { text: '02', rtl: true },
      { text: '-', rtl: true },
      { text: '2026', rtl: true },
    ];
    const { order } = computeLineVisualOrder(segs, true);
    const visual = order.map((i) => segs[i].text).join('');
    expect(visual).toBe('28-02-2026');
  });
});
