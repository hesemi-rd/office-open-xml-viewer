import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData, TabStop } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.2.7 (a:pPr@rtl) + §21.1.2.1.x (a:tabLst / a:tab @pos @algn) +
// UAX#9 (bidi reordering) — a right-/centre-aligned tab stop in a RIGHT-TO-LEFT
// paragraph must resolve in the RTL READING FRAME, not the LTR absolute one.
//
// The bug (issue #831, latent — no RTL+tabLst fixture shipped): the tab-stop
// draw path anchored the stop at `bx + lPad + pos` (LTR: measured LEFTWARD from
// the LEFT text-inset edge) and placed the trailing cell with LTR math
// (`tabAbsX − totalTabW` for @algn="r"). Under an RTL base a tab advances the pen
// in READING order — from the LEADING (right) inset edge, moving LEFT — so the
// stop sits at `(bx + bw − rIns) − pos` and the cell mirrors: an @algn="r" (end)
// cell puts its TRAILING (left) edge on the stop. The LTR path dropped the cell
// on the wrong visual side. This mirrors the docx fix (#830 / #835:
// `layoutBidiTabStops` / `nextTabStopRtl`, resolving stops against the leading
// text margin in the reading frame). DrawingML tabs carry no leader, so unlike
// docx there is nothing to paint across the gap.
//
// The mock ctx measures every glyph at FONT_PX (no contextual collapse here — we
// use plain Latin/digits), so widths are exact and the cell X is deterministic.

const FONT_PX = 20;
const SCALE = 1 / 12700; // emuToPx(emu, SCALE) = emu·SCALE; PT_TO_EMU=12700 ⇒ 1pt → 1px

function mockCtx(): {
  ctx: CanvasRenderingContext2D;
  texts: { text: string; x: number; direction: CanvasDirection }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const texts: { text: string; x: number; direction: CanvasDirection }[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string, x: number) => texts.push({ text: t, x, direction }),
    strokeText: () => {},
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function run(text: string, color = '000000'): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PX, color, fontFamily: 'Serif',
  } as TextRunData;
}

/** A paragraph carrying `tabStops`. The run text contains a `\t` so layout
 *  emits an inline tab segment (#916 multi-cell model).
 *
 *  `alignment` follows the PARSER's resolution order (parser/src/text.rs:617):
 *  explicit `pPr@algn` → inherited body/layout/master default (not modelled
 *  here) → fallback `"r"` when `rtl="1"`, else `"l"`. Pass `opts.algn` to model
 *  an EXPLICIT paragraph alignment; omit it for the parser fallback. Tab-stop
 *  mirroring must follow the BASE DIRECTION (`rtl`), never the visual
 *  alignment. `opts.marR` (EMU) is the paragraph's logical-leading (physical
 *  RIGHT) margin under an RTL base. */
function bodyWithTab(
  tabStops: TabStop[],
  runs: TextRunData[],
  opts: { rtl?: boolean; algn?: string; marR?: number } = {},
): TextBody {
  const rtl = opts.rtl ?? false;
  const para: Paragraph = {
    alignment: opts.algn ?? (rtl ? 'r' : 'l'),
    marL: 0, marR: opts.marR ?? 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops, rtl, runs,
  } as unknown as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: FONT_PX,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
}

type RunInfo = { text: string; inShapeX: number; w: number };

function render(body: TextBody, boxW = 600): {
  texts: { text: string; x: number; direction: CanvasDirection }[];
  runs: RunInfo[];
} {
  const { ctx, texts } = mockCtx();
  const runs: RunInfo[] = [];
  renderTextBody(
    ctx, body, 0, 0, boxW, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    (r) => runs.push({ text: r.text, inShapeX: r.inShapeX, w: r.w }),
  );
  return { texts, runs };
}

// Box 600px wide, no insets/margins. Tab stop 400px from the LEADING inset edge.
// LTR leading edge = left = 0 ⇒ stop canvas X = 0 + 400 = 400.
// RTL leading edge = right = boxW − rIns = 600 ⇒ stop canvas X = 600 − 400 = 200.
const BOX_W = 600;
const TAB_POS_PX = 400;
const TAB_POS_EMU = TAB_POS_PX * 12700;
const CELL = '12'; // 2 glyphs → 40px
const CELL_W = [...CELL].length * FONT_PX;
const LTR_STOP_X = TAB_POS_PX; // 400
const RTL_STOP_X = BOX_W - TAB_POS_PX; // 200

describe('pptx RTL tab stops resolve in the reading frame (issue #831, mirrors docx #830/#835)', () => {
  // (1) @algn="r" (end) in an RTL paragraph: the cell's TRAILING (left) edge lands
  //     ON the mirrored stop and the cell extends rightward. The LTR path put the
  //     cell's RIGHT edge on the LTR stop (x = 400 − 40 = 360) — the wrong side.
  it('places a right-aligned tab cell at the mirrored stop (trailing edge on the stop)', () => {
    const { runs } = render(
      bodyWithTab([{ pos: TAB_POS_EMU, algn: 'r' }], [run(`\t${CELL}`)], { rtl: true }),
      BOX_W,
    );
    const cell = runs.find((r) => r.text === CELL)!;
    expect(cell, 'cell run reported').toBeTruthy();
    // RTL: trailing(left) edge on the stop ⇒ pen = stopX = 200.
    expect(cell.inShapeX).toBeCloseTo(RTL_STOP_X, 6);
    // Guard against the LTR-frame regression (360).
    expect(cell.inShapeX).not.toBeCloseTo(LTR_STOP_X - CELL_W, 3);
  });

  // (2) @algn="ctr" in an RTL paragraph: the cell centres on the mirrored stop.
  it('centres a centre-aligned tab cell on the mirrored stop', () => {
    const { runs } = render(
      bodyWithTab([{ pos: TAB_POS_EMU, algn: 'ctr' }], [run(`\t${CELL}`)], { rtl: true }),
      BOX_W,
    );
    const cell = runs.find((r) => r.text === CELL)!;
    expect(cell.inShapeX).toBeCloseTo(RTL_STOP_X - CELL_W / 2, 6); // 200 − 20 = 180
  });

  // (3) ITEM-2 CORRECTION (issue #916): tab-cell content now flows through the
  //     per-segment UAX#9 pass (`computeLineVisualOrder`) instead of a blanket
  //     `ctx.direction='rtl'` for the whole cell. A DIGITS cell ("12") is
  //     Bidi_Class EN → even (LTR) level even under an RTL base, so it is drawn
  //     with ctx.direction='ltr' (digits read left-to-right in Arabic too). The
  //     old model's unconditional 'rtl' was the item-2 defect; mixed-direction
  //     cell reordering is pinned by bidi-line.test.ts / pptx-multi-tab.test.ts.
  it('draws a digits tab cell with ctx.direction = ltr (per-segment bidi, item 2)', () => {
    const { texts } = render(
      bodyWithTab([{ pos: TAB_POS_EMU, algn: 'r' }], [run(`\t${CELL}`)], { rtl: true }),
      BOX_W,
    );
    const cellDraw = texts.find((t) => t.text === CELL)!;
    expect(cellDraw.direction).toBe('ltr');
  });

  // (4) An EXPLICIT `pPr@algn="l"` on an rtl="1" paragraph (the parser resolution
  //     keeps the explicit value; the r-fallback applies only when algn is absent)
  //     must NOT turn off the mirroring: the stop frame follows the BASE
  //     DIRECTION, not the visual alignment.
  it('mirrors the stop under rtl even when the paragraph alignment is explicitly "l"', () => {
    const { runs } = render(
      bodyWithTab([{ pos: TAB_POS_EMU, algn: 'r' }], [run(`\t${CELL}`)], { rtl: true, algn: 'l' }),
      BOX_W,
    );
    const cell = runs.find((r) => r.text === CELL)!;
    expect(cell.inShapeX).toBeCloseTo(RTL_STOP_X, 6);
  });

  // (5) A start ('l') tab advances the reading-frame pen so a later end stop stays
  //     reachable. Stops: start@100px, end@140px; marR=50px; 'A'=20px. Reading
  //     frame (pen from the leading indent = marR): pen 50→(start tab)→100, +A→120
  //     ⇒ the end stop at 140 is still ahead ⇒ the '\t12' cell EXISTS. This is the
  //     reachability property case (5) has always verified.
  //
  //     ITEM-3 CORRECTION (issue #916): the start-tab gap is now MATERIALISED
  //     (previously it silently advanced the pen without rendering), so the line
  //     is laid out cumulatively. The end stop at reading 140 is BEHIND the pen
  //     (120) once its 40-wide cell is placed (target 140−40 = 100 < 120), so the
  //     end tab collapses per the no-backward-tab clamp — the same behaviour as
  //     docx `layoutBidiTabStops`. The cell's trailing (left) edge therefore lands
  //     at the cumulative pen, not the absolute mirrored stop: content width = 50
  //     (start gap) + 20 (A) + 0 (collapsed end tab) + 40 (cell) = 110, so the
  //     RTL-anchored cell sits at (600 − marR − 110) = 440, with 'A' at 480 (the
  //     rendered start gap pulls it off the leading edge — the visible item-3 fix).
  it('advances a start tab in the reading frame so a later end stop stays reachable', () => {
    const stops: TabStop[] = [
      { pos: 100 * 12700, algn: 'l' },
      { pos: 140 * 12700, algn: 'r' },
    ];
    const { runs } = render(
      bodyWithTab(stops, [run(`\tA\t${CELL}`)], { rtl: true, marR: 50 * 12700 }),
      BOX_W,
    );
    const cell = runs.find((r) => r.text === CELL);
    expect(cell, 'end-stop cell exists (start-tab advance did not overshoot)').toBeTruthy();
    expect(cell!.inShapeX).toBeCloseTo(440, 6);
    // The start-tab gap is now rendered: 'A' is pulled off the leading (right)
    // text edge (was flush at 530 in the pre-fix inline-advance model).
    const a = runs.find((r) => r.text === 'A')!;
    expect(a.inShapeX).toBeCloseTo(480, 6);
  });

  // (6) A stop past the TRAILING (left) text edge pins the cell at that edge —
  //     the docx layoutBidiTabStops clamp (#835: Word never pushes the cell off
  //     the text area). Unclamped, the mirrored stop 600−700 = −100 would draw
  //     the cell at a NEGATIVE x, outside the shape.
  it('pins the cell at the trailing text edge when the stop overflows the text area', () => {
    const { runs } = render(
      bodyWithTab([{ pos: 700 * 12700, algn: 'r' }], [run(`\t${CELL}`)], { rtl: true }),
      BOX_W,
    );
    const cell = runs.find((r) => r.text === CELL)!;
    expect(cell, 'cell run reported').toBeTruthy();
    expect(cell.inShapeX).toBeCloseTo(0, 6); // bx + lPad = 0 — never negative
  });

  // (7) REGRESSION GUARD — the LTR path is byte-identical: an LTR right tab still
  //     anchors at the LTR stop with the cell's RIGHT edge on it (x = 400 − 40).
  it('leaves the LTR right-tab path unchanged (cell right edge on the LTR stop)', () => {
    const { texts, runs } = render(
      bodyWithTab([{ pos: TAB_POS_EMU, algn: 'r' }], [run(`\t${CELL}`)]),
      BOX_W,
    );
    const cell = runs.find((r) => r.text === CELL)!;
    expect(cell.inShapeX).toBeCloseTo(LTR_STOP_X - CELL_W, 6); // 360
    const cellDraw = texts.find((t) => t.text === CELL)!;
    expect(cellDraw.direction).toBe('ltr');
  });
});
