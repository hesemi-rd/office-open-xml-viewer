import { describe, it, expect } from 'vitest';
import { computeFloatTableBox, registerTableFloat, floatTableWrapSide } from './float-table-geometry.js';
import type { FloatTableBox } from './float-table-geometry.js';
import type { TblpPr } from './types.js';
import type { FloatRect } from './float-layout.js';

// Table-driven geometry assertions for floating tables (ECMA-376 §17.4.57
// `<w:tblpPr>` / §17.4.56 `<w:tblOverlap>`). The VRT does not cover the
// floating-table sample (private/sample-11 is not in the VRT suite), so these
// pin the placement math AND the FloatRect that registerTableFloat emits
// (xLeft/xRight/yTop/yBottom/mode/side), which resolveLineFloatWindow consumes
// to wrap the surrounding body text.
//
// Geometry is exercised at scale=1 so px == pt. A representative page:
//   pageWidth=600, margins L/R/T/B = 100/100/72/72 ⇒ content band [100,500].
//   A single column is modeled as contentX=100, contentW=400; hAnchor="text"
//   snaps to it (the #513 column-relative contract).

interface MinState {
  scale: number;
  contentX: number;
  contentW: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  pageWidth: number;
  pageH: number;
  floats: FloatRect[];
  floatParaSeq: number;
}

function makeState(over: Partial<MinState> = {}): MinState {
  return {
    scale: 1,
    contentX: 100,
    contentW: 400,
    marginLeft: 100,
    marginRight: 100,
    marginTop: 72,
    marginBottom: 72,
    pageWidth: 600,
    pageH: 800,
    floats: [],
    floatParaSeq: 0,
    ...over,
  };
}

// Full TblpPr with the spec defaults; tests override only the axis under test.
// The factory sets horzAnchor:'page' (an explicit horizontal hint), so
// horzSpecified defaults to true; the "no horizontal spec at all" case is
// exercised by overriding horzSpecified:false (see the sample-11 case).
function tblp(over: Partial<TblpPr> = {}): TblpPr {
  return {
    leftFromText: 0,
    rightFromText: 0,
    topFromText: 0,
    bottomFromText: 0,
    horzAnchor: 'page',
    horzSpecified: true,
    vertAnchor: 'page',
    tblpX: 0,
    tblpY: 0,
    ...over,
  };
}

// Cast helpers: the geometry fns read only the MinState subset of RenderState.
const box = (tp: TblpPr, st: MinState, paraTop: number, w: number, h: number): FloatTableBox =>
  computeFloatTableBox(tp, st as never, paraTop, w, h);
const registerFloat = (
  b: FloatTableBox,
  tp: TblpPr,
  st: MinState,
  side: string,
  allowOverlap: boolean,
): void => registerTableFloat(b, tp, st as never, side, allowOverlap);
const wrapSide = (b: FloatTableBox, st: MinState): string =>
  floatTableWrapSide(b, st as never);

describe('floating table geometry (§17.4.57) — horizontal placement', () => {
  it('horzAnchor="page", tblpX=0 ⇒ table at the physical page left edge', () => {
    const st = makeState();
    const b = box(tblp({ horzAnchor: 'page', tblpX: 0 }), st, 300, 150, 60);
    expect(b.x).toBe(0); // page left + 0
    expect(b.w).toBe(150);
    expect(b.h).toBe(60);
  });

  it('horzAnchor="page", tblpX signed offset is added (signed twips round-trip)', () => {
    const st = makeState();
    expect(box(tblp({ horzAnchor: 'page', tblpX: 36 }), st, 300, 150, 60).x).toBe(36);
    expect(box(tblp({ horzAnchor: 'page', tblpX: -20 }), st, 300, 150, 60).x).toBe(-20);
  });

  it('horzAnchor="margin", tblpX anchors against the page content margin', () => {
    const st = makeState();
    const b = box(tblp({ horzAnchor: 'margin', tblpX: 10 }), st, 300, 150, 60);
    expect(b.x).toBe(100 + 10); // marginLeft + tblpX
  });

  it('horzAnchor="text", tblpX anchors against the COLUMN band (contentX, #513)', () => {
    const st = makeState({ contentX: 250, contentW: 200 }); // a right-hand column
    const b = box(tblp({ horzAnchor: 'text', tblpX: 10 }), st, 300, 80, 40);
    expect(b.x).toBe(250 + 10); // column left + tblpX
  });

  it('tblpXSpec="left" / "inside" snap to the band left, superseding tblpX', () => {
    const st = makeState();
    for (const spec of ['left', 'inside']) {
      const b = box(tblp({ horzAnchor: 'text', tblpX: 999, tblpXSpec: spec }), st, 300, 100, 40);
      expect(b.x, spec).toBe(100); // contentX, tblpX ignored
    }
  });

  it('tblpXSpec="center" centers the table in the band, superseding tblpX', () => {
    const st = makeState(); // text band [100,500], width 400
    const b = box(tblp({ horzAnchor: 'text', tblpX: 999, tblpXSpec: 'center' }), st, 300, 100, 40);
    expect(b.x).toBe(100 + (400 - 100) / 2);
  });

  it('tblpXSpec="right" / "outside" right-align the table in the band', () => {
    const st = makeState();
    for (const spec of ['right', 'outside']) {
      const b = box(tblp({ horzAnchor: 'text', tblpXSpec: spec }), st, 300, 100, 40);
      expect(b.x, spec).toBe(500 - 100); // band right − tableW
    }
  });
});

describe('floating table geometry (§17.4.57) — vertical placement', () => {
  it('vertAnchor="text", tblpY ⇒ table top at the anchor paragraph top + tblpY', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'text', tblpY: 1 }), st, 300, 150, 60);
    expect(b.y).toBe(300 + 1); // paraTop + tblpY
  });

  it('vertAnchor="margin", tblpY anchors against the top content margin', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 5 }), st, 300, 150, 60);
    expect(b.y).toBe(72 + 5); // marginTop + tblpY
  });

  it('vertAnchor="page", tblpY anchors against the physical page top', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'page', tblpY: 5 }), st, 300, 150, 60);
    expect(b.y).toBe(0 + 5); // page top + tblpY
  });

  it('vertAnchor="text" IGNORES tblpYSpec (relative vertical not allowed) → tblpY', () => {
    const st = makeState();
    // tblpYSpec="center" would otherwise center; with vertAnchor="text" it is
    // ignored and the absolute tblpY offset is used instead (§17.4.57).
    const b = box(tblp({ vertAnchor: 'text', tblpY: 7, tblpYSpec: 'center' }), st, 300, 150, 60);
    expect(b.y).toBe(300 + 7);
  });

  it('vertAnchor="margin" + tblpYSpec="bottom" bottom-aligns on the page, superseding tblpY', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 999, tblpYSpec: 'bottom' }), st, 300, 150, 60);
    expect(b.y).toBe((800 - 72) - 60); // (pageH − marginBottom) − tableH
  });

  it('vertAnchor="margin" + tblpYSpec="center" centers vertically, superseding tblpY', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 999, tblpYSpec: 'center' }), st, 300, 150, 60);
    // margin band [72, 728]: start + (end − start − h)/2 = 72 + (728 − 72 − 60)/2
    expect(b.y).toBe(72 + (728 - 72 - 60) / 2);
  });

  it('vertAnchor="margin" + tblpYSpec="center": ASYMMETRIC margins center in the margin band, NOT the page', () => {
    // §22.9.2.20: tblpYSpec="center" centers in the margin BAND
    // [marginTop, pageH−marginBottom], which only equals the page centre when
    // margins are symmetric. The symmetric case above (marginTop=marginBottom=72)
    // cannot distinguish "margin band centre" from "page centre"; this asymmetric
    // case pins the band-relative behaviour.
    const st = makeState({ marginTop: 40, marginBottom: 120 }); // band [40, 680], height 640
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 999, tblpYSpec: 'center' }), st, 300, 150, 60);
    // band centre = marginTop + (pageH − marginTop − marginBottom − h)/2
    //             = 40 + (800 − 40 − 120 − 60)/2 = 40 + 290 = 330
    expect(b.y).toBe(40 + (800 - 40 - 120 - 60) / 2); // 330 — NOT page centre (800−60)/2 = 370
    expect(b.y).not.toBe((800 - 60) / 2); // explicit: this is the band centre, not the page centre
  });

  // §22.9.2.20: ST_YAlign positions relative to the ANCHOR OBJECT (the vertAnchor
  // band), NOT the physical page. For vertAnchor="page" the band is [0, pageH], so
  // center/bottom must NOT carry a margin offset (the pre-band code subtracted
  // marginTop even for page-anchored tables, §17.18.100 page = page edges).
  it('vertAnchor="page" + tblpYSpec="center" centers over the FULL page (no margin offset)', () => {
    const st = makeState(); // page band [0, 800]
    const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: 'center' }), st, 300, 150, 60);
    expect(b.y).toBe((800 - 60) / 2); // 370 — NOT (800−60)/2 − marginTop
  });

  it('vertAnchor="page" + tblpYSpec="bottom" sits flush to the physical page bottom', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: 'bottom' }), st, 300, 150, 60);
    expect(b.y).toBe(800 - 60); // pageH − tableH (no marginBottom inset)
  });

  it('vertAnchor="page" + tblpYSpec="outside" sits flush to the physical page bottom', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: 'outside' }), st, 300, 150, 60);
    expect(b.y).toBe(800 - 60);
  });

  it('vertAnchor="page" + tblpYSpec="top"/"inside" sits at the page top (band start)', () => {
    const st = makeState();
    for (const spec of ['top', 'inside']) {
      const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: spec }), st, 300, 150, 60);
      expect(b.y, spec).toBe(0); // page band start = 0
    }
  });
});

describe('floating table float registration (§17.4.57 / §17.4.56)', () => {
  it('registers a square wrap float padded by the *FromText dist values', () => {
    const st = makeState();
    const tp = tblp({ horzAnchor: 'page', leftFromText: 5, rightFromText: 9, topFromText: 3, bottomFromText: 7 });
    const b = box(tp, st, 300, 150, 60);
    registerFloat(b, tp, st, 'right', true);
    expect(st.floats).toHaveLength(1);
    const f = st.floats[0];
    expect(f.mode).toBe('square');
    expect(f.side).toBe('right');
    expect(f.imageKey).toBe(''); // non-image float (table painted directly)
    expect(f.drawn).toBe(true);
    expect(f.xLeft).toBe(0 - 5); // box.x − leftFromText
    expect(f.xRight).toBe(0 + 150 + 9); // box.x + tableW + rightFromText
    expect(f.yTop).toBe(b.y - 3);
    expect(f.yBottom).toBe(b.y + 60 + 7);
    expect(f.distLeft).toBe(5);
    expect(f.distRight).toBe(9);
  });

  it('a degenerate (zero-area) box registers no float', () => {
    const st = makeState();
    const tp = tblp();
    registerFloat({ x: 0, y: 0, w: 0, h: 0 }, tp, st, 'right', true);
    expect(st.floats).toHaveLength(0);
  });

  it('overlap="never" (allowOverlap=false) re-seats the new float off another table', () => {
    // Pre-seat a blocking FLOATING TABLE [0,200]×[300,360] from another paragraph.
    const st = makeState({
      floatParaSeq: 1,
      floats: [{
        kind: 'table', mode: 'square', imageKey: '', imageX: 0, imageY: 300, imageW: 200, imageH: 60,
        xLeft: 0, xRight: 200, yTop: 300, yBottom: 360, side: 'bothSides',
        distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0, drawn: true,
      }],
    });
    // A new floating table at page-left x=0 overlapping the blocker; never ⇒ avoid.
    const tp = tblp({ horzAnchor: 'page', tblpX: 0, vertAnchor: 'page', tblpY: 320 });
    const b = box(tp, st, 0, 100, 20); // box at (0,320), overlaps the blocker
    registerFloat(b, tp, st, 'right', /* allowOverlap */ false);
    const f = st.floats[st.floats.length - 1];
    // resolveFloatOverlap re-seats it to the right of the blocker's right edge.
    expect(f.xLeft).toBeGreaterThanOrEqual(200);
  });

  // §17.4.56 scope: a floating table with <w:tblOverlap w:val="never"/> must
  // only avoid OTHER FLOATING TABLES. DrawingML anchors (§20.4.2.3) and text
  // frames keep their own placement — a never-overlap table does NOT re-seat off
  // them. These pin that scoping via FloatRect.kind.
  describe('overlap="never" scopes to other floating tables only (§17.4.56)', () => {
    // A blocker occupying [0,200]×[300,360], anchored in another paragraph,
    // parameterized by kind so we can assert table-vs-non-table behavior.
    const blocker = (kind: FloatRect['kind']): FloatRect => ({
      kind, mode: 'square', imageKey: '', imageX: 0, imageY: 300, imageW: 200, imageH: 60,
      xLeft: 0, xRight: 200, yTop: 300, yBottom: 360, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0, drawn: true,
    });
    // A never-overlap floating table seated at page-left x=0, overlapping [0,200].
    const seatNeverTable = (blockers: FloatRect[]): FloatRect => {
      const st = makeState({ floatParaSeq: 1, floats: blockers });
      const tp = tblp({ horzAnchor: 'page', tblpX: 0, vertAnchor: 'page', tblpY: 320 });
      const b = box(tp, st, 0, 100, 20); // box at (0,320), overlaps [0,200]
      registerFloat(b, tp, st, 'right', /* allowOverlap (never) */ false);
      return st.floats[st.floats.length - 1];
    };

    it('does NOT re-seat off a DrawingML shape/image float (stays at x=0)', () => {
      const f = seatNeverTable([blocker('shape')]);
      expect(f.xLeft).toBe(0); // overlap permitted: §17.4.56 ignores non-tables
    });

    it('does NOT re-seat off a text-frame float (stays at x=0)', () => {
      const f = seatNeverTable([blocker('frame')]);
      expect(f.xLeft).toBe(0);
    });

    it('DOES re-seat off another floating-table float', () => {
      const f = seatNeverTable([blocker('table')]);
      expect(f.xLeft).toBeGreaterThanOrEqual(200); // avoids the other table
    });

    it('avoids only the table blocker when a shape and a table both overlap', () => {
      // shape at [0,200], table at [0,200] (same band, different paragraphs):
      // the table re-seats past the TABLE's right edge, ignoring the shape.
      const f = seatNeverTable([blocker('shape'), blocker('table')]);
      expect(f.xLeft).toBeGreaterThanOrEqual(200);
    });
  });
});

describe('floating table — sample-11 case (§17.4.57)', () => {
  // <w:tblpPr w:rightFromText="187" w:bottomFromText="72" w:vertAnchor="text"
  //           w:tblpY="1"/> + <w:tblOverlap w:val="never"/>. NO horzAnchor, NO
  // tblpX, NO tblpXSpec ⇒ horzSpecified=false. The spec-literal default would be
  // horzAnchor="page" + tblpX=0 (page edge), but Word anchors such a table at the
  // anchor paragraph's text/column LEFT (contentX), not the page edge — matching
  // the calibre PDF where the ITEM/NEEDED table aligns with the body column.
  // rightFromText 187 twip → 9.35 pt; bottomFromText 72 twip → 3.6 pt.
  const RIGHT_FROM_TEXT = 187 / 20; // 9.35 pt
  const BOTTOM_FROM_TEXT = 72 / 20; // 3.6 pt

  it('no horizontal spec ⇒ anchors at the text/column left (NOT the page edge)', () => {
    const st = makeState(); // column band [100, 500]
    const tp = tblp({
      rightFromText: RIGHT_FROM_TEXT,
      bottomFromText: BOTTOM_FROM_TEXT,
      horzAnchor: 'page', // defaulted token, but…
      horzSpecified: false, // …no horizontal hint was present in the source
      vertAnchor: 'text',
      tblpX: 0,
      tblpY: 1,
    });
    // A small ITEM/NEEDED table, say 120 wide × 50 tall, anchored at paraTop=300.
    const tableW = 120;
    const b = box(tp, st, 300, tableW, 50);
    expect(b.x).toBe(100); // text/column left (contentX), NOT page-left 0
    expect(b.y).toBe(300 + 1); // vAnchor=text ⇒ paraTop + tblpY

    // side from the column band [100,500]: the float's right edge (100+120=220)
    // is left of the column centre (300) ⇒ float on the LEFT ⇒ text wraps RIGHT.
    const side = wrapSide(b, st);
    expect(side).toBe('right'); // text sits to the RIGHT of the float

    registerFloat(b, tp, st, side, /* overlap="never" ⇒ */ false);
    const f = st.floats[0];
    expect(f.side).toBe('right');
    expect(f.xRight).toBe(100 + tableW + RIGHT_FROM_TEXT); // right edge includes the dist padding
    expect(f.yBottom).toBe(b.y + 50 + BOTTOM_FROM_TEXT);
  });

  it('explicit horzAnchor="page" + tblpX=0 still pins to the page edge', () => {
    const st = makeState();
    // When the source DOES specify a horizontal anchor, honor the spec literally.
    const b = box(tblp({ horzAnchor: 'page', horzSpecified: true, tblpX: 0 }), st, 300, 120, 50);
    expect(b.x).toBe(0); // page left edge
  });
});

describe('floating table wrap side (§17.4.57) from the column band', () => {
  const st = makeState(); // column band [100, 500], centre 300

  it('float whose right edge is left-of-centre ⇒ text wraps RIGHT', () => {
    expect(wrapSide({ x: 0, y: 0, w: 120, h: 50 }, st)).toBe('right');
  });

  it('float whose left edge is right-of-centre ⇒ text wraps LEFT', () => {
    expect(wrapSide({ x: 350, y: 0, w: 120, h: 50 }, st)).toBe('left');
  });

  it('float straddling the column centre ⇒ bothSides', () => {
    expect(wrapSide({ x: 200, y: 0, w: 200, h: 50 }, st)).toBe('bothSides');
  });
});
