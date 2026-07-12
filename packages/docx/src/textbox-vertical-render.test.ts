import { describe, it, expect } from 'vitest';
import { measureShapeTextAutoFitHeight, renderShapeText } from './renderer';
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

interface ImageCall {
  /** Net rotation of the drawn image in device space, degrees (atan2(b,a)). */
  angleDeg: number;
  /** Device-space position of the draw origin (local (x,y) mapped by the CTM). */
  devX: number;
  devY: number;
  /** Local draw width/height (before the CTM). */
  w: number;
  h: number;
}

/** Recording 2D context that tracks the full affine CTM (a,b,c,d,e,f) across
 *  save/restore/translate/rotate/scale, so a glyph's net orientation is
 *  recoverable at fillText time. measureText gives every code point the current
 *  font px as advance (1 em / CJK), with symmetric font/ink boxes. */
function makeMatrixCtx(): {
  ctx: CanvasRenderingContext2D;
  glyphs: GlyphCall[];
  images: ImageCall[];
} {
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
  const images: ImageCall[] = [];
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
    drawImage(_bmp: unknown, x: number, y: number, w: number, h: number) {
      const angleDeg = (Math.atan2(m.b, m.a) * 180) / Math.PI;
      const devX = m.a * x + m.c * y + m.e;
      const devY = m.b * x + m.d * y + m.f;
      images.push({ angleDeg, devX, devY, w, h });
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, glyphs, images };
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

  // ── (a) mongolianVert (§20.1.10.83) ──────────────────────────────────────
  // GT (batch-3 adjudication): identical per-glyph orientation to eaVert (CJK
  // UPRIGHT, Latin sideways 90° CW), but the line/column progression is the
  // MIRROR of eaVert — columns advance LEFT→RIGHT instead of right→left.
  it('mongolianVert: CJK upright (0°) while Latin stays sideways (+90°) — same as eaVert', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    renderShapeText(richTextbox([run(CJK), run(LAT)], 'mongolianVert'), 0, 0, 200, 100, ctx, 1, {});
    const cjk = glyphs.find((g) => g.text.includes(CJK));
    const lat = glyphs.find((g) => g.text.includes(LAT));
    expect(cjk, 'CJK glyph drawn').toBeDefined();
    expect(lat, 'Latin glyph drawn').toBeDefined();
    expect(NEAR(norm(cjk!.angleDeg), 0), `CJK @${cjk!.angleDeg}`).toBe(true);
    expect(NEAR(norm(lat!.angleDeg), 90), `Latin @${lat!.angleDeg}`).toBe(true);
  });

  it('mongolianVert stacks columns LEFT→RIGHT (mirror of eaVert R→L)', () => {
    // Two blocks (two columns). eaVert puts the first column on the RIGHT and the
    // second to its left; mongolianVert mirrors it — first column on the LEFT,
    // second to its right.
    const twoCols = (v: string) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const block2: ShapeText = { text: CJK, fontSizePt: 10, alignment: 'left', runs: [run(CJK)] };
      const shape = richTextbox([run(CJK)], v);
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(block2);
      renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
      return glyphs;
    };
    const ea = twoCols('eaVert');
    const mn = twoCols('mongolianVert');
    // eaVert: first column right of the last. mongolianVert: first column LEFT.
    expect(ea[0].devX).toBeGreaterThan(ea[ea.length - 1].devX);
    expect(mn[0].devX, `mongolianVert first col @${mn[0].devX} < last @${mn[mn.length - 1].devX}`)
      .toBeLessThan(mn[mn.length - 1].devX);
  });

  // ── (b) eaVert + ruby (§17.3.3.25) ───────────────────────────────────────
  // GT: furigana sits on the RIGHT side of the vertical base column, upright,
  // running top→bottom. In the +90° CW frame the physical RIGHT is device +x.
  it('eaVert ruby draws furigana upright on the device-RIGHT of the base column', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    const baseRun: ShapeTextRun = {
      text: '漢字',
      fontSizePt: 10,
      fontFamily: 'NotInMetrics',
      ruby: { text: 'かんじ', fontSizePt: 5 },
    };
    renderShapeText(richTextbox([baseRun], 'eaVert'), 0, 0, 200, 100, ctx, 1, {});
    const base = glyphs.filter((g) => /[漢字]/.test(g.text));
    const ruby = glyphs.filter((g) => /[かんじ]/.test(g.text));
    expect(base.length, 'base glyphs drawn').toBeGreaterThan(0);
    expect(ruby.length, 'ruby glyphs drawn').toBe(3);
    // Ruby stands upright (net 0°), like the upright CJK base cells.
    for (const r of ruby) expect(NEAR(norm(r.angleDeg), 0), `ruby ${r.text}@${r.angleDeg}`).toBe(true);
    // Physical RIGHT = device +x: every ruby glyph is right of every base glyph.
    const baseMaxX = Math.max(...base.map((g) => g.devX));
    const rubyMinX = Math.min(...ruby.map((g) => g.devX));
    expect(rubyMinX, `ruby minX ${rubyMinX} > base maxX ${baseMaxX}`).toBeGreaterThan(baseMaxX);
    // Exact cross offset: base fontSize 10 (cell centred on the column baseline),
    // ruby fontSize 5 ⇒ ruby column centre = baseline − (effSize/2 + rubySize) =
    // baseline − 10, which maps to device +10 (the physical right). Base and ruby
    // sit on constant device-x per column, so the mean difference isolates the
    // cross offset.
    const meanBaseX = base.reduce((s, g) => s + g.devX, 0) / base.length;
    const meanRubyX = ruby.reduce((s, g) => s + g.devX, 0) / ruby.length;
    expect(meanRubyX - meanBaseX, `cross offset ${meanRubyX - meanBaseX} ≈ effSize/2+rubySize=10`)
      .toBeCloseTo(10, 0);
  });

  // ── (c) vert + inline image: image stays UPRIGHT ─────────────────────────
  // GT: the inline raster keeps its physical orientation (a graphic is not text),
  // even though the surrounding `vert` glyphs rotate 90° CW.
  it('vert inline image is drawn UPRIGHT (net 0°), not rotated with the text frame', () => {
    const { ctx, images } = makeMatrixCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 24, imageHeightPt: 36,
    } as unknown as ShapeText;
    const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'vert');
    (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(imgBlock);
    const fakeBmp = { width: 24, height: 36 } as unknown as ImageBitmap;
    const imgs = new Map<string, ImageBitmap>([['word/media/image1.png', fakeBmp]]);
    renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {}, imgs as never);
    expect(images.length, 'one image drawn').toBe(1);
    // Upright: the net rotation cancels the +90° page frame → 0°.
    expect(NEAR(norm(images[0].angleDeg), 0), `image @${images[0].angleDeg}`).toBe(true);
    // Portrait aspect preserved (physical 24×36, not swapped to 36×24). The draw
    // dimensions in the upright local frame are width=physical-width,
    // height=physical-height (the callback receives dw=cross, dh=along).
    expect(Math.abs(images[0].w)).toBeCloseTo(24, 3);
    expect(Math.abs(images[0].h)).toBeCloseTo(36, 3);
  });

  it('vert270 inline image is ALSO drawn upright (net 0°), not −180° (counter-rotation sign)', () => {
    // vert270's page frame is −90°, so the image counter-rotation must be +90°
    // (not the −90° the +90° modes use) — otherwise the raster is flipped 180°.
    const { ctx, images } = makeMatrixCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 24, imageHeightPt: 36,
    } as unknown as ShapeText;
    const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'vert270');
    (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(imgBlock);
    const fakeBmp = { width: 24, height: 36 } as unknown as ImageBitmap;
    const imgs = new Map<string, ImageBitmap>([['word/media/image1.png', fakeBmp]]);
    renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {}, imgs as never);
    expect(images.length).toBe(1);
    expect(NEAR(norm(images[0].angleDeg), 0), `vert270 image @${images[0].angleDeg}`).toBe(true);
  });

  it('vertical inline image reserves its physical WIDTH (crossExtent), not its height, along the column stack', () => {
    // Two CJK text columns sandwich the image column; their cross (device-x)
    // separation = text-line-box + the image's reserved cross extent. A tall-thin
    // image (10 wide × 90 tall) reserves its physical WIDTH 10 (crossExtent) — the
    // fitH-vs-crossExtent bug would instead reserve its 90-tall height, pushing the
    // trailing column ~80px further, so the separation cleanly distinguishes them.
    const { ctx, glyphs } = makeMatrixCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 10, imageHeightPt: 90,
    } as unknown as ShapeText;
    const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'vert');
    const tb = shape as unknown as { textBlocks: ShapeText[] };
    tb.textBlocks.push(imgBlock);
    tb.textBlocks.push({ text: CJK, fontSizePt: 10, alignment: 'left', runs: [run(CJK)] } as ShapeText);
    const fakeBmp = { width: 10, height: 90 } as unknown as ImageBitmap;
    const imgs = new Map<string, ImageBitmap>([['word/media/image1.png', fakeBmp]]);
    renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {}, imgs as never);
    const cjk = glyphs.filter((g) => g.text.includes(CJK));
    const firstCol = cjk[0].devX;
    const lastCol = cjk[cjk.length - 1].devX;
    const sep = Math.abs(firstCol - lastCol);
    // Separation ≈ one text line box (~10-14) + image cross 10 ≈ 20-30. With the
    // bug (reserving the 90 height) it would exceed 90. Assert well below 90.
    expect(sep, `two-column separation ${sep} reflects image cross 10, not height 90`).toBeLessThan(45);
    expect(sep, `columns are actually separated ${sep}`).toBeGreaterThan(12);
  });

  // ── cross-axis spacing (sample-53 class regressions) ─────────────────────
  // GT (Word PDF): a ruby-bearing line's COLUMN is widened by the ruby
  // reservation (the same `ruby.fontSizePt × 1.5` ascent bump layoutLines()
  // already computes for the section body), so the base column clears the
  // previous column instead of the furigana overprinting it.
  it('eaVert ruby line widens its column advance by the ruby reservation (§17.3.3.25)', () => {
    const advance = (withRuby: boolean): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const labelBlock: ShapeText = {
        text: CJK, fontSizePt: 10, alignment: 'left',
        runs: [{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }],
      };
      const baseRun: ShapeTextRun = {
        text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics',
        ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 5 } } : {}),
      };
      const baseBlock: ShapeText = { text: '漢', fontSizePt: 10, alignment: 'left', runs: [baseRun] };
      const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'eaVert');
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks = [labelBlock, baseBlock];
      renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
      const label = glyphs.find((g) => g.text.includes(CJK));
      const base = glyphs.find((g) => g.text.includes('漢'));
      expect(label, 'label glyph drawn').toBeDefined();
      expect(base, 'base glyph drawn').toBeDefined();
      // eaVert stacks columns R→L: the base column sits LEFT of the label column.
      return label!.devX - base!.devX;
    };
    const plain = advance(false);
    const withRuby = advance(true);
    // layoutLines() reserves ruby.fontSizePt × 1.5 = 7.5 of extra ascent; the
    // ruby-bearing column must advance by exactly that much more.
    expect(withRuby - plain, `ruby advance ${withRuby} vs plain ${plain}`).toBeCloseTo(7.5, 5);
  });

  // GT (Word PDF): mongolianVert's first column is the exact MIRROR of the
  // eaVert first column (sample-53: flow-start-edge → label glyph-box gap is
  // preserved within 0.5pt). Reflecting only the band ORIGIN keeps the baseline
  // measured from the band's physical RIGHT edge, so asymmetric leading (line
  // box taller than the glyph em box) shoves the mongolian line toward the
  // physical left — the centerline inside the band must be reflected too.
  it('mongolianVert mirrors the line centerline within its cross-axis band', () => {
    const firstColumnCenterX = (mode: 'eaVert' | 'mongolianVert'): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const shape = richTextbox([{ text: CJK, fontSizePt: 11, fontFamily: 'NotInMetrics' }], mode);
      const [block] = (shape as unknown as { textBlocks: ShapeText[] }).textBlocks;
      // Exact 20pt line box > the 11pt glyph box → asymmetric leading exposes a
      // non-mirrored centerline (natural == lineH would accidentally pass).
      (block as unknown as { lineSpacingRule: string; lineSpacingVal: number }).lineSpacingRule = 'exact';
      (block as unknown as { lineSpacingRule: string; lineSpacingVal: number }).lineSpacingVal = 20;
      renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
      const glyph = glyphs.find((g) => g.text.includes(CJK));
      expect(glyph, 'first-column glyph drawn').toBeDefined();
      // Upright CJK cells draw centred on the column centerline, so devX IS the
      // centerline's cross position.
      return glyph!.devX;
    };
    const eaX = firstColumnCenterX('eaVert');
    const mongolianX = firstColumnCenterX('mongolianVert');
    // Equal glyphs have equal cross extents, so mirrored centerlines imply equal
    // flow-start-edge → glyph-box gaps (box spans x ∈ [0, 200], insets 0).
    expect(mongolianX, `mongolian centerline ${mongolianX} mirrors eaVert ${eaX}`).toBeCloseTo(200 - eaX, 5);
  });

  // mongolianVert + ruby: the furigana keeps its orientation side (physical
  // RIGHT of the base column, like eaVert) — so the mirror must keep the BASE
  // cell at its non-ruby mirrored position and spend the ruby reservation on
  // the band's interior (right) side. Mirroring the ruby-inflated baseline
  // offset naively would shove the base toward the band's right edge and paint
  // the furigana on top of the NEXT column.
  it('mongolianVert ruby line keeps its base column mirrored and reserves toward the next column', () => {
    const render = (withRuby: boolean) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const rubyBlock: ShapeText = {
        text: '漢', fontSizePt: 10, alignment: 'left',
        runs: [{
          text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics',
          ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 5 } } : {}),
        }],
      };
      const nextBlock: ShapeText = {
        text: CJK, fontSizePt: 10, alignment: 'left',
        runs: [{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }],
      };
      const shape = richTextbox([{ text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'mongolianVert');
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks = [rubyBlock, nextBlock];
      renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
      const base = glyphs.find((g) => g.text.includes('漢'));
      const next = glyphs.find((g) => g.text.includes(CJK));
      const ruby = glyphs.filter((g) => /[かん]/.test(g.text));
      expect(base, 'base glyph drawn').toBeDefined();
      expect(next, 'next-column glyph drawn').toBeDefined();
      if (withRuby) expect(ruby.length, 'ruby glyphs drawn').toBe(2);
      return { baseX: base!.devX, nextX: next!.devX, rubyXs: ruby.map((g) => g.devX) };
    };
    const plain = render(false);
    const withRuby = render(true);
    // The base column does NOT move: the reservation goes to the ruby side
    // (band interior), not the flow-start edge.
    expect(withRuby.baseX, `base ${withRuby.baseX} unmoved from ${plain.baseX}`).toBeCloseTo(plain.baseX, 5);
    // The NEXT column (to the right, L→R) advances by the ruby reservation.
    expect(withRuby.nextX - plain.nextX, 'next column pushed by the reservation').toBeCloseTo(7.5, 5);
    // The furigana sits BETWEEN its base column and the next column — physical
    // right of the base (orientation preserved), clear of the next column.
    for (const rx of withRuby.rubyXs) {
      expect(rx, `ruby ${rx} right of base ${withRuby.baseX}`).toBeGreaterThan(withRuby.baseX);
      expect(rx, `ruby ${rx} clears next column ${withRuby.nextX}`).toBeLessThan(withRuby.nextX - 5);
    }
  });

  // Autofit (spAutoFit) shares the same line resolver: a ruby-bearing vertical
  // line must widen the fitted cross extent by the same reservation, or the
  // grown box would clip the furigana (sample-53 box (d) class).
  it('eaVert spAutoFit measurement grows by the ruby reservation', () => {
    const measure = (withRuby: boolean): number => {
      const { ctx } = makeMatrixCtx();
      const shape = richTextbox([{
        text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics',
        ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 5 } } : {}),
      }], 'eaVert');
      return measureShapeTextAutoFitHeight(shape, 200, ctx, 1, {});
    };
    expect(measure(true) - measure(false), 'fitted extent grows by ruby.fontSizePt × 1.5').toBeCloseTo(7.5, 5);
  });

  // GT (Word PDF): §21.1.2.1.1 names the insets by PHYSICAL bounding-box edge
  // (lIns = left inset, default 91440 EMU = 7.2pt; bIns = bottom, 45720 EMU) —
  // reflecting the column stack for mongolianVert's L→R flow must not hand the
  // physical LEFT edge to bIns.
  it('mongolianVert first column origin honors lIns (physical left), not bIns (§21.1.2.1.1)', () => {
    const firstColX = (lIns: number, bIns: number): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'mongolianVert');
      const s = shape as unknown as {
        textInsetL: number; textInsetT: number; textInsetR: number; textInsetB: number;
      };
      s.textInsetL = lIns;
      s.textInsetT = 3;
      s.textInsetR = 7;
      s.textInsetB = bIns;
      renderShapeText(shape, 0, 0, 200, 100, ctx, 1, {});
      const g = glyphs.find((gl) => gl.text.includes(CJK));
      expect(g, 'glyph drawn').toBeDefined();
      return g!.devX;
    };
    // +12pt of lIns moves the first (leftmost) column +12pt right.
    expect(firstColX(17, 3) - firstColX(5, 3), 'lIns owns the physical-left origin').toBeCloseTo(12, 5);
    // bIns must NOT move the physical-left first column.
    expect(firstColX(7, 20), 'bIns does not own the physical-left origin').toBeCloseTo(firstColX(7, 3), 5);
  });
});
