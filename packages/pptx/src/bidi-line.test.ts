import { describe, it, expect } from 'vitest';
import { segmentsHaveRtl, computeLineVisualOrder } from './bidi-line.js';

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

  // ── TAB segments are Bidi_Class S (§17.3.1.37 / UAX#9) — issue #916 item 2 ──
  // A tab is a Segment Separator, NOT a neutral object. Rules L1/L2 reset it to
  // the paragraph level and reorder each tab-delimited CELL independently, so an
  // RTL paragraph's tab-aligned cells appear in mirrored (leading-cell-at-the-
  // right) order. Mirrors docx bidi-line.ts.
  it('reorders LTR cells in mirrored order under an RTL base (tab = S)', () => {
    // logical [A][TAB][B][TAB][C], base RTL → visual [C, TAB, B, TAB, A].
    const segs = [
      { text: 'A' },
      { isTab: true },
      { text: 'B' },
      { isTab: true },
      { text: 'C' },
    ];
    const { order } = computeLineVisualOrder(segs, true);
    expect(order.map((i) => ('isTab' in segs[i] ? 'TAB' : segs[i].text))).toEqual([
      'C', 'TAB', 'B', 'TAB', 'A',
    ]);
  });

  it('keeps tab-delimited LTR cells in logical order under an LTR base', () => {
    const segs = [{ text: 'A' }, { isTab: true }, { text: 'B' }];
    const { order } = computeLineVisualOrder(segs, false);
    expect(order).toEqual([0, 1, 2]);
  });

  it('reorders MIXED-direction content INSIDE a tab cell (item 2)', () => {
    // A cell (after a tab) holding a Hebrew word then a Latin word under an RTL
    // base: the Latin word is visually LEFT of the Hebrew word within the cell.
    // Old model drew the cell sequentially in logical order (no reorder).
    const segs = [{ isTab: true }, { text: 'אב' }, { text: 'CD' }];
    const { order, rtl } = computeLineVisualOrder(segs, true);
    // visual L→R: [CD, אב, TAB]
    expect(order.map((i) => ('isTab' in segs[i] ? 'TAB' : segs[i].text))).toEqual([
      'CD', 'אב', 'TAB',
    ]);
    // the Latin cell content resolves LTR, the Hebrew RTL.
    const hebIdx = 1;
    const latIdx = 2;
    expect(rtl[hebIdx]).toBe(true);
    expect(rtl[latIdx]).toBe(false);
  });
});
