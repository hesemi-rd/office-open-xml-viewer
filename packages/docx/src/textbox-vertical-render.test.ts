import { describe, it, expect } from 'vitest';
import { renderShapeText } from './renderer';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types';

// ECMA-376 §20.1.10.83 ST_TextVerticalType — a DrawingML text-box body direction
// (`<wps:bodyPr vert>`). Word writes it on a Word text box (`<wps:txbx>` shape):
//   - `vert`    : ALL glyphs rotated 90° CW  (chars T→B, lines R→L).
//   - `vert270` : ALL glyphs rotated 270° CW (= 90° CCW; chars B→T, lines L→R).
//   - `eaVert`  : East-Asian upright vertical — CJK stands UPRIGHT, non-EA glyphs
//                 rotated 90° (chars T→B, lines R→L). Mirrors the section-level
//                 tbRl per-glyph path (UAX#50 vo) and pptx's verified eaVert.
//   - `horz` / absent : horizontal (unchanged legacy path).
//
// The renderer laies the box out with the SAME horizontal engine, rotated ±90°
// about the box centre with width/height swapped, so the layout/kinsoku/bidi are
// reused. Because the text box does NOT emit `onTextRun`, we verify by recording
// the Canvas transform at every glyph draw and asserting each glyph's NET
// rotation (the angle its local +x axis makes in device space).

interface GlyphCall {
  text: string;
  /** Net rotation of the drawn glyph in device space, degrees (atan2(b,a)). */
  angleDeg: number;
  /** Device-space position of the draw origin (local (x,y) mapped by the CTM). */
  devX: number;
  devY: number;
}

/** Recording 2D context that tracks the full affine CTM (a,b,c,d,e,f) across
 *  save/restore/translate/rotate/scale, so a glyph's net orientation is
 *  recoverable at fillText time. measureText gives every code point the current
 *  font px as advance (1 em / CJK), with symmetric font/ink boxes. */
function makeMatrixCtx(): { ctx: CanvasRenderingContext2D; glyphs: GlyphCall[] } {
  let m = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const stack: (typeof m)[] = [];
  let font = '10px serif';
  let textAlign = 'start';
  let textBaseline = 'alphabetic';
  let letterSpacing = '0px';
  let fillStyle = '#000';
  let direction = 'ltr';
  let fontKerning = 'auto';
  const glyphs: GlyphCall[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get textAlign() { return textAlign; },
    set textAlign(v: string) { textAlign = v; },
    get textBaseline() { return textBaseline; },
    set textBaseline(v: string) { textBaseline = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; },
    set direction(v: string) { direction = v; },
    get fontKerning() { return fontKerning; },
    set fontKerning(v: string) { fontKerning = v; },
    strokeStyle: '#000',
    lineWidth: 1,
    globalAlpha: 1,
    save() { stack.push({ ...m }); },
    restore() { const s = stack.pop(); if (s) m = s; },
    translate(tx: number, ty: number) {
      m = { ...m, e: m.e + m.a * tx + m.c * ty, f: m.f + m.b * tx + m.d * ty };
    },
    rotate(t: number) {
      const cos = Math.cos(t), sin = Math.sin(t);
      m = {
        a: m.a * cos + m.c * sin,
        b: m.b * cos + m.d * sin,
        c: -m.a * sin + m.c * cos,
        d: -m.b * sin + m.d * cos,
        e: m.e,
        f: m.f,
      };
    },
    scale(sx: number, sy: number) {
      m = { ...m, a: m.a * sx, b: m.b * sx, c: m.c * sy, d: m.d * sy };
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, rect() {},
    fill() {}, stroke() {}, clip() {}, fillRect() {}, strokeRect() {}, clearRect() {},
    setTransform() {}, resetTransform() {},
    measureText(s: string) {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) {
      const angleDeg = (Math.atan2(m.b, m.a) * 180) / Math.PI;
      const devX = m.a * x + m.c * y + m.e;
      const devY = m.b * x + m.d * y + m.f;
      glyphs.push({ text, angleDeg, devX, devY });
    },
    strokeText() {},
    drawImage() {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, glyphs };
}

function richTextbox(
  runs: ShapeTextRun[],
  textVert?: string | null,
  alignment = 'left',
): ShapeRun {
  const block: ShapeText = {
    text: runs.map((r) => r.text).join(''),
    fontSizePt: runs[0]?.fontSizePt ?? 10,
    alignment,
    runs,
  };
  return {
    type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textVert: textVert ?? null,
  } as unknown as ShapeRun;
}

const CJK = '経'; // UAX#50 vo=U (upright)
const LAT = 'A'; // vo=R (sideways)
const NEAR = (a: number, b: number, tol = 1e-6) => Math.abs(a - b) <= tol;
// Normalise an angle to (-180, 180].
const norm = (deg: number) => ((((deg + 180) % 360) + 360) % 360) - 180;

describe('§20.1.10.83 textbox <wps:bodyPr vert> — vertical text-box rendering', () => {
  const run = (text: string): ShapeTextRun => ({ text, fontSizePt: 10, fontFamily: 'NotInMetrics' });

  it('horz (absent vert): no rotation — glyphs upright, CTM identity (legacy path)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    renderShapeText(richTextbox([run(CJK + LAT)]), 0, 0, 200, 100, ctx, 1, {});
    const drawn = glyphs.filter((g) => g.text.includes(CJK) || g.text.includes(LAT) || /経|A/.test(g.text));
    expect(drawn.length).toBeGreaterThan(0);
    for (const g of drawn) expect(NEAR(norm(g.angleDeg), 0)).toBe(true);
  });

  it('vert: every glyph rotated +90° CW (all-rotate; CJK included)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    renderShapeText(richTextbox([run(CJK + LAT)], 'vert'), 0, 0, 200, 100, ctx, 1, {});
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) expect(NEAR(norm(g.angleDeg), 90), `${g.text}@${g.angleDeg}`).toBe(true);
  });

  it('vert270: every glyph rotated −90° (270° CW)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    renderShapeText(richTextbox([run(CJK + LAT)], 'vert270'), 0, 0, 200, 100, ctx, 1, {});
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) expect(NEAR(norm(g.angleDeg), -90), `${g.text}@${g.angleDeg}`).toBe(true);
  });

  it('eaVert: CJK stands UPRIGHT (net 0°) while Latin stays sideways (net +90°)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    renderShapeText(richTextbox([run(CJK), run(LAT)], 'eaVert'), 0, 0, 200, 100, ctx, 1, {});
    const cjk = glyphs.find((g) => g.text.includes(CJK));
    const lat = glyphs.find((g) => g.text.includes(LAT));
    expect(cjk, 'CJK glyph drawn').toBeDefined();
    expect(lat, 'Latin glyph drawn').toBeDefined();
    expect(NEAR(norm(cjk!.angleDeg), 0), `CJK @${cjk!.angleDeg}`).toBe(true);
    expect(NEAR(norm(lat!.angleDeg), 90), `Latin @${lat!.angleDeg}`).toBe(true);
  });

  it('rotated glyphs land INSIDE the physical box (transform pivots on box centre)', () => {
    // A vert box of physical 200×100: after the +90° rotation about the centre,
    // every drawn glyph's device origin must still fall within the physical box.
    const { ctx, glyphs } = makeMatrixCtx();
    renderShapeText(richTextbox([run(CJK + CJK + LAT)], 'vert'), 0, 0, 200, 100, ctx, 1, {});
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) {
      expect(g.devX, `${g.text} devX in [0,200]`).toBeGreaterThanOrEqual(-1);
      expect(g.devX).toBeLessThanOrEqual(201);
      expect(g.devY, `${g.text} devY in [0,100]`).toBeGreaterThanOrEqual(-1);
      expect(g.devY).toBeLessThanOrEqual(101);
    }
  });

  it('eaVert justified (both) column advances by NATURAL width — no §17.18.44 stretch drift', () => {
    // A justified (`both`) eaVert paragraph that WRAPS: the first column has slack
    // (a long Latin word can't break mid-word, so it wraps whole), which the
    // horizontal justify pass would distribute into inter-segment gaps. Inside an
    // eaVert cell that distribution must NOT be painted — the column flows
    // start-aligned by its natural measured widths — so the four CJK cells on the
    // first column stay UNIFORMLY spaced across the two-run boundary. (The prior
    // bug advanced by the justify-expanded width, opening a gap at the run seam.)
    const { ctx, glyphs } = makeMatrixCtx();
    const two: ShapeTextRun[] = [run('経経'), run('済済'), run('ABCDEFGHIJ')];
    const block: ShapeText = {
      text: '経経済済ABCDEFGHIJ',
      fontSizePt: 10,
      alignment: 'both',
      runs: two,
    };
    const shape = {
      type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
      textBlocks: [block], textAnchor: 't',
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textVert: 'eaVert',
    } as unknown as ShapeRun;
    // Box 200×100 → logical column length 100 → 10 cells of the 10px font. The
    // first column holds 経経済済 (4 cells); ABCDEFGHIJ (10 cells) wraps whole.
    renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
    const cjk = glyphs.filter((g) => /[経済]/.test(g.text) && [...g.text].length === 1);
    expect(cjk.length, 'four upright CJK cells on the justified first column').toBe(4);
    // The along-column axis is device +y (the +90° frame maps local +x → +y).
    const dys = cjk.map((g) => g.devY).sort((a, b) => a - b);
    const gaps = dys.slice(1).map((v, i) => v - dys[i]);
    const minGap = Math.min(...gaps);
    const maxGap = Math.max(...gaps);
    // All three inter-cell gaps equal (uniform natural pitch); the bug widened the
    // run1→run2 gap by the distributed slack.
    expect(maxGap - minGap, `uniform cell pitch, gaps=${gaps}`).toBeLessThan(0.5);
  });

  it('vert vs vert270 stack lines to OPPOSITE sides of the box centre', () => {
    // Two lines (a hard wrap via two blocks) → the second line sits on the
    // opposite cross-side for vert (R→L, leftwards) vs vert270 (L→R, rightwards).
    const two = (v: string) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const block2: ShapeText = { text: CJK, fontSizePt: 10, alignment: 'left', runs: [run(CJK)] };
      const shape = richTextbox([run(CJK)], v);
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(block2);
      renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
      return glyphs;
    };
    const gv = two('vert');
    const g270 = two('vert270');
    // Box centre is at device x = 100 (w=200). vert: lines go R→L, so line 1 sits
    // right of centre, line 2 left. vert270: L→R, mirrored.
    expect(gv[0].devX).toBeGreaterThan(gv[gv.length - 1].devX);
    expect(g270[0].devX).toBeLessThan(g270[g270.length - 1].devX);
  });
});
