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
});
