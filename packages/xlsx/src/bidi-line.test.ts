import { describe, it, expect } from 'vitest';
import { segmentsHaveRtl, computeLineVisualOrder, cellBaseRtl } from './bidi-line.js';

describe('cellBaseRtl', () => {
  it('honors explicit readingOrder', () => {
    expect(cellBaseRtl(2, 'Hello')).toBe(true); // 2 = RTL
    expect(cellBaseRtl(1, 'مرحبا')).toBe(false); // 1 = LTR
  });
  it('uses first-strong for Context (0) / absent', () => {
    expect(cellBaseRtl(0, 'مرحبا')).toBe(true);
    expect(cellBaseRtl(0, 'Hello')).toBe(false);
    expect(cellBaseRtl(undefined, 'منتج alpha')).toBe(true);
    expect(cellBaseRtl(undefined, '1234')).toBe(false);
  });
});

describe('segmentsHaveRtl', () => {
  it('detects Arabic/Hebrew, ignores Latin and objects', () => {
    expect(segmentsHaveRtl([{ text: 'Hello world' }])).toBe(false);
    expect(segmentsHaveRtl([{ text: 'price 99' }, {}])).toBe(false); // object seg has no text
    expect(segmentsHaveRtl([{ text: 'مرحبا' }])).toBe(true);
    expect(segmentsHaveRtl([{ text: 'abc ' }, { text: 'שלום' }])).toBe(true);
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

  it('places a neutral object by surrounding direction', () => {
    // object between two Hebrew words, base RTL -> object stays in RTL flow
    const { order } = computeLineVisualOrder([{ text: 'שלום ' }, {}, { text: 'טוב' }], true);
    expect(order).toEqual([2, 1, 0]);
  });
});
