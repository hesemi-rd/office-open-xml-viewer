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

/** A paragraph carrying `tabStops`. The run text must contain a `\t` so layout
 *  switches into tab-stop accumulation mode.
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

  // (3) RTL cell glyphs are painted with ctx.direction = 'rtl' so bidi content in
  //     the cell shapes in reading order (textAlign stays 'left', so the pen X is
  //     unchanged — only glyph shaping/joining differs).
  it('draws the RTL tab cell with ctx.direction = rtl', () => {
    const { texts } = render(
      bodyWithTab([{ pos: TAB_POS_EMU, algn: 'r' }], [run(`\t${CELL}`)], { rtl: true }),
      BOX_W,
    );
    const cellDraw = texts.find((t) => t.text === CELL)!;
    expect(cellDraw.direction).toBe('rtl');
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

  // (5) A start ('l') tab opens NO cell — it advances the PEN inline (the gap is
  //     not rendered; pre-existing model, both directions). Under an RTL base the
  //     advance must live in the READING frame (distance from the leading indent,
  //     = marR), like the stop selection: mixing frames (selection from marR,
  //     advance from marL) overshoots the pen when marR > marL and a later stop
  //     becomes unreachable. Stops: start@100px, end@140px; marR=50px; 'A'=20px.
  //     Reading frame: pen 50→(start tab)→100, +A→120 → the end stop at 140 is
  //     still ahead ⇒ '\t12' opens the cell, trailing edge on 600−140=460.
  //     Mixed-frame (the bug): lineW=100−marL(0)=100, +A→120 → selection pen
  //     50+120=170 > 140 ⇒ no stop ⇒ the tab degrades to a space (no cell).
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
    expect(cell!.inShapeX).toBeCloseTo(BOX_W - 140, 6); // 460
  });

  // (6) REGRESSION GUARD — the LTR path is byte-identical: an LTR right tab still
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
