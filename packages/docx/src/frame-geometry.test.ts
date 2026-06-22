import { describe, it, expect } from 'vitest';
import { computeFrameBox, registerFrameFloat } from './frame-geometry.js';
import type { FrameBox } from './frame-geometry.js';
import type { FramePr } from './types.js';
import type { FloatRect } from './float-layout.js';

// Table-driven geometry assertions for paragraph frames / drop caps
// (ECMA-376 §17.3.1.11). The VRT only exercises dropCap="drop" wrap="around"
// (private/sample-11 page 5), so the other (wrap, dropCap, hAnchor, vAnchor)
// combinations are pinned here — this is their only regression guard. Each case
// asserts the resolved FrameBox AND the FloatRect that registerFrameFloat emits
// (xLeft/xRight/yTop/yBottom/mode/side), which is what resolveLineFloatWindow
// consumes to wrap the following body text.
//
// Geometry is exercised at scale=1 so px == pt. A representative page:
//   pageWidth=600, margins L/R/T/B = 100/100/72/72  ⇒ content band [100,500].
//   A multi-column run would set a narrower contentX/contentW; we model a single
//   column here as contentX=100, contentW=400 and assert hAnchor="text" snaps to
//   it (the #513 column-relative contract).

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

// Full FramePr with the spec defaults; tests override only the axis under test.
function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none',
    lines: 1,
    wrap: 'around',
    hAnchor: 'text',
    vAnchor: 'text',
    hRule: 'auto',
    hSpace: 0,
    vSpace: 0,
    ...over,
  };
}

// Cast helper: computeFrameBox/registerFrameFloat read only the MinState subset
// of RenderState exercised here.
const box = (fp: FramePr, st: MinState, paraTop: number, cW: number, cH: number, anchorH: number): FrameBox =>
  computeFrameBox(fp, st as never, paraTop, cW, cH, anchorH);
const registerFloat = (b: FrameBox, fp: FramePr, st: MinState): void =>
  registerFrameFloat(b, fp, st as never);

describe('frame geometry (§17.3.1.11) — drop cap placement', () => {
  it('dropCap="drop": frame at column left, height = lines × anchor line height', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'drop', lines: 3, hAnchor: 'text', vAnchor: 'text' });
    // paraTop=200 (in-flow Y), measured cap width 42, anchor line height 14.
    const b = box(fp, st, 200, 42, 50, 14);
    expect(b.x).toBe(100); // column left (contentX)
    expect(b.y).toBe(200); // paragraph top (vAnchor="text")
    expect(b.w).toBe(42); // auto width = measured cap advance
    expect(b.h).toBe(3 * 14); // lines × anchor line height (y/yAlign ignored)
  });

  it('dropCap="margin": frame hangs into the left margin (left = band left − width)', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'margin', lines: 2, hAnchor: 'text' });
    const b = box(fp, st, 150, 30, 40, 12);
    expect(b.x).toBe(100 - 30); // outside the column margin
    expect(b.w).toBe(30);
    expect(b.h).toBe(2 * 12);
  });

  it('drop cap emits a right-side square float (text wraps to the right only)', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'drop', lines: 3, wrap: 'around', hSpace: 8 });
    const b = box(fp, st, 200, 42, 50, 14);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    const f = st.floats[0];
    expect(f.mode).toBe('square');
    expect(f.side).toBe('right');
    // hSpace=8 pads L/R for wrap="around".
    expect(f.xLeft).toBe(100 - 8);
    expect(f.xRight).toBe(100 + 42 + 8);
    expect(f.yTop).toBe(200);
    expect(f.yBottom).toBe(200 + 42); // h=3×14, vSpace=0
  });
});

describe('frame geometry (§17.3.1.11) — wrap modes', () => {
  const st0 = () => makeState();
  const dc = (wrap: FramePr['wrap']) => frame({ dropCap: 'drop', lines: 3, wrap });

  it('wrap="notBeside" → topAndBottom float (text never beside the frame)', () => {
    const st = st0();
    const fp = dc('notBeside');
    const b = box(fp, st, 200, 42, 50, 14);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    expect(st.floats[0].mode).toBe('topAndBottom');
  });

  it('wrap="around" and "auto" → square float (auto ≡ around in Word)', () => {
    for (const w of ['around', 'auto'] as const) {
      const st = st0();
      const fp = dc(w);
      const b = box(fp, st, 200, 42, 50, 14);
      registerFloat(b, fp, st);
      expect(st.floats, `wrap=${w}`).toHaveLength(1);
      expect(st.floats[0].mode, `wrap=${w}`).toBe('square');
    }
  });

  it('wrap="tight" and "through" → square float (rectangle, no contour follow)', () => {
    for (const w of ['tight', 'through'] as const) {
      const st = st0();
      const fp = dc(w);
      const b = box(fp, st, 200, 42, 50, 14);
      registerFloat(b, fp, st);
      expect(st.floats, `wrap=${w}`).toHaveLength(1);
      expect(st.floats[0].mode, `wrap=${w}`).toBe('square');
    }
  });

  it('wrap="none" → no float registered (absolute draw only, no exclusion)', () => {
    const st = st0();
    const fp = dc('none');
    const b = box(fp, st, 200, 42, 50, 14);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(0);
  });
});

describe('frame geometry (§17.3.1.11) — hAnchor / vAnchor containers', () => {
  // A generic (non-drop-cap) frame at an absolute x/y from each anchor base.
  it('hAnchor="text" anchors x against the COLUMN band (contentX)', () => {
    const st = makeState({ contentX: 250, contentW: 200 }); // a right-hand column
    const fp = frame({ hAnchor: 'text', x: 10 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.x).toBe(250 + 10); // column left + x offset
  });

  it('hAnchor="margin" anchors x against the page content margin', () => {
    const st = makeState();
    const fp = frame({ hAnchor: 'margin', x: 10 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.x).toBe(100 + 10); // marginLeft + x
  });

  it('hAnchor="page" anchors x against the physical page edge', () => {
    const st = makeState();
    const fp = frame({ hAnchor: 'page', x: 10 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.x).toBe(0 + 10); // page left + x
  });

  it('vAnchor="text" anchors y at the paragraph top', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'text', y: 5 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.y).toBe(300 + 5); // paraTop + y (yAlign ignored when vAnchor=text)
  });

  it('vAnchor="margin" anchors y at the top content margin', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', y: 5 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.y).toBe(72 + 5); // marginTop + y
  });

  it('vAnchor="page" anchors y at the physical page top', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'page', y: 5 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.y).toBe(0 + 5); // page top + y
  });
});

describe('frame geometry (§17.3.1.11 / §22.9.2.20) — yAlign is vAnchor-band relative', () => {
  // ST_YAlign positions the frame relative to the ANCHOR OBJECT (the vAnchor
  // band), NOT the physical page (§22.9.2.20: "relative position … relative to
  // the vertical anchor"). The band per §17.18.100:
  //   page   → [0, pageH]                       (page edges)
  //   margin → [marginTop, pageH−marginBottom]  (text margins)
  //   text   → [paraTop, paraTop+contentH]      (anchor paragraph text extents)
  // yAlign is ignored for vAnchor="text" (relative positioning not allowed).

  it('vAnchor="margin" + yAlign="center" centers in the MARGIN band, not the page', () => {
    const st = makeState(); // margin band [72, 728], height 656
    const fp = frame({ vAnchor: 'margin', yAlign: 'center', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    // start + (end − start − h)/2 = 72 + (728 − 72 − 60)/2 = 72 + 298 = 370
    expect(b.y).toBe(72 + (728 - 72 - 60) / 2);
  });

  it('vAnchor="margin" + yAlign="center": ASYMMETRIC margins center in the margin band, NOT the page', () => {
    // §22.9.2.20: with vAnchor="margin", yAlign="center" centers in the margin
    // BAND [marginTop, pageH−marginBottom], which only equals the page centre
    // when margins are symmetric. The symmetric cases above (marginTop=
    // marginBottom=72) cannot distinguish "margin band centre" from "page
    // centre"; this asymmetric case pins the band-relative behaviour.
    const st = makeState({ marginTop: 40, marginBottom: 120 }); // band [40, 680], height 640
    const fp = frame({ vAnchor: 'margin', yAlign: 'center', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    // band centre = marginTop + (pageH − marginTop − marginBottom − h)/2
    //             = 40 + (800 − 40 − 120 − 60)/2 = 40 + 290 = 330
    expect(b.y).toBe(40 + (800 - 40 - 120 - 60) / 2); // 330 — NOT page centre (800−60)/2 = 370
    expect(b.y).not.toBe((800 - 60) / 2); // explicit: this is the band centre, not the page centre
  });

  it('vAnchor="margin" + yAlign="bottom" sits flush to the bottom margin', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', yAlign: 'bottom', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(728 - 60); // (pageH − marginBottom) − h
  });

  it('vAnchor="margin" + yAlign="outside" sits flush to the bottom margin', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', yAlign: 'outside', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(728 - 60);
  });

  it('vAnchor="margin" + yAlign="top"/"inside" sits at the margin top (band start)', () => {
    const st = makeState();
    for (const ya of ['top', 'inside'] as const) {
      const fp = frame({ vAnchor: 'margin', yAlign: ya, hRule: 'exact', h: 60 });
      const b = box(fp, st, 300, 40, 100, 12);
      expect(b.y, ya).toBe(72); // band start = marginTop
    }
  });

  it('vAnchor="page" + yAlign="center" centers over the FULL page (no margin offset)', () => {
    const st = makeState(); // page band [0, 800]
    const fp = frame({ vAnchor: 'page', yAlign: 'center', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe((800 - 60) / 2); // 370 — NOT (800−60)/2 − marginTop
  });

  it('vAnchor="page" + yAlign="bottom" sits flush to the physical page bottom', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'page', yAlign: 'bottom', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(800 - 60); // pageH − h
  });

  it('explicit y is measured from the vAnchor band start (margin top)', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', y: 5 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(72 + 5); // band start (marginTop) + y
  });

  it('yAlign is ignored for vAnchor="text" (relative positioning not allowed)', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'text', yAlign: 'center', y: 7 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(300 + 7); // paraTop + y; yAlign ignored
  });
});

describe('frame geometry (§17.3.1.11) — generic frame (dropCap="none") sizing', () => {
  it('hRule="exact" forces the frame height to h regardless of content', () => {
    const st = makeState();
    const fp = frame({ hRule: 'exact', h: 30 });
    const b = box(fp, st, 200, 60, 100, 12); // contentH 100 ignored
    expect(b.h).toBe(30);
  });

  it('hRule="atLeast" takes max(h, content height)', () => {
    const st = makeState();
    expect(box(frame({ hRule: 'atLeast', h: 30 }), st, 200, 60, 100, 12).h).toBe(100);
    expect(box(frame({ hRule: 'atLeast', h: 200 }), st, 200, 60, 100, 12).h).toBe(200);
  });

  it('hRule="auto" uses the content height; explicit w forces exact width', () => {
    const st = makeState();
    const b = box(frame({ hRule: 'auto', w: 150 }), st, 200, 60, 80, 12);
    expect(b.h).toBe(80); // content height
    expect(b.w).toBe(150); // explicit width supersedes natural content width
  });

  it('xAlign="center" centers the frame in the hAnchor band, superseding x', () => {
    const st = makeState(); // text band [100,500], width 400
    const fp = frame({ hAnchor: 'text', x: 999, xAlign: 'center', w: 100 });
    const b = box(fp, st, 200, 100, 40, 12);
    expect(b.x).toBe(100 + (400 - 100) / 2); // centered, x ignored
  });

  it('xAlign="right" right-aligns the frame in the hAnchor band', () => {
    const st = makeState();
    const fp = frame({ hAnchor: 'text', xAlign: 'right', w: 100 });
    const b = box(fp, st, 200, 100, 40, 12);
    expect(b.x).toBe(500 - 100); // band right − width
  });

  it('a generic frame emits a bothSides square float (text wraps either side)', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'none', hRule: 'exact', h: 40, w: 120, x: 50 });
    const b = box(fp, st, 200, 120, 40, 12);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    expect(st.floats[0].side).toBe('bothSides');
  });
});
