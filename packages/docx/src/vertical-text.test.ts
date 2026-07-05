import { describe, it, expect } from 'vitest';
import {
  isUprightVerticalGlyph,
  verticalDrawMode,
  verticalGlyphOffset,
  splitVerticalOrientationRuns,
  drawVerticalRun,
  drawTateChuYokoRun,
  drawUprightBox,
  physicalToLogicalAnchorBox,
  verticalTextLayerPlacement,
} from './vertical-text.js';

// ECMA-376 §17.6.20 vertical writing (tbRl). These are the pure classification
// + geometry primitives the renderer wires into the glyph/image draw path behind
// the `verticalCJK` flag. Orientation is decided by the Unicode UAX#50
// Vertical_Orientation property (core `verticalOrientation`).

describe('verticalDrawMode (UAX#50 vo → draw mode)', () => {
  const cp = (ch: string): number => ch.codePointAt(0) ?? 0;

  it('U/Tu → upright (ideographs, kana, 、。！？, small kana)', () => {
    for (const ch of ['富', '士', 'あ', 'ア', '、', '。', '！', '？', 'ぁ', 'ッ']) {
      expect(verticalDrawMode(cp(ch))).toBe('upright');
    }
  });

  it('Tr → rotate (long vowel mark ー, corner brackets, parens, quotes)', () => {
    for (const ch of ['ー', '「', '」', '（', '）', '〈', '〉', '“', '”']) {
      expect(verticalDrawMode(cp(ch))).toBe('rotate');
    }
  });

  it('R → sideways (Latin, digits, ASCII punctuation)', () => {
    for (const ch of ['A', 'z', '0', '5', '9', '@', '-', '.']) {
      expect(verticalDrawMode(cp(ch))).toBe('sideways');
    }
  });
});

describe('isUprightVerticalGlyph (UAX#50 vo ∈ {U, Tu})', () => {
  const cp = (ch: string): number => ch.codePointAt(0) ?? 0;

  it('is true for U/Tu (ideographs, kana, 、。) and false for Tr/R', () => {
    for (const ch of ['富', 'あ', '、', '。']) expect(isUprightVerticalGlyph(cp(ch))).toBe(true);
    // ー「」（） are Tr (rotate), Latin/digits are R — not upright.
    for (const ch of ['ー', '「', '）', 'A', '5']) expect(isUprightVerticalGlyph(cp(ch))).toBe(false);
  });
});

describe('verticalGlyphOffset (upper-right nudge — fallback when no vertical form)', () => {
  const cp = (ch: string): number => ch.codePointAt(0) ?? 0;

  it('nudges ． (FF0E, no U+FExx vertical form) toward the upper-right corner', () => {
    const off = verticalGlyphOffset(cp('．'));
    expect(off.dx).toBeGreaterThan(0); // rightward
    expect(off.dy).toBeLessThan(0); // upward
  });

  it('returns {0,0} for glyphs that get a substituted vertical form (、。，) or need no shift', () => {
    // These have vertical presentation forms → substituted, not nudged.
    for (const ch of ['、', '。', '，', '富', 'A', 'ー']) {
      expect(verticalGlyphOffset(cp(ch))).toEqual({ dx: 0, dy: 0 });
    }
  });
});

describe('splitVerticalOrientationRuns (§17.6.20 — group by draw mode)', () => {
  it('splits a mixed run into maximal same-mode pieces in logical order', () => {
    const pieces = splitVerticalOrientationRuns('第5回大会');
    expect(pieces).toEqual([
      { text: '第', mode: 'upright' },
      { text: '5', mode: 'sideways' },
      { text: '回大会', mode: 'upright' },
    ]);
  });

  it('separates a Tr bracket/長音符 into its own rotate piece', () => {
    // チーム(土): チ=upright, ー=rotate(Tr), ム=upright, (=rotate(Tr), 土=upright, )=rotate(Tr)
    const pieces = splitVerticalOrientationRuns('チーム（土）');
    expect(pieces).toEqual([
      { text: 'チ', mode: 'upright' },
      { text: 'ー', mode: 'rotate' },
      { text: 'ム', mode: 'upright' },
      { text: '（', mode: 'rotate' },
      { text: '土', mode: 'upright' },
      { text: '）', mode: 'rotate' },
    ]);
  });

  it('keeps a pure-CJK run as one upright piece', () => {
    expect(splitVerticalOrientationRuns('富士町')).toEqual([{ text: '富士町', mode: 'upright' }]);
  });

  it('keeps a pure-Latin run as one sideways piece', () => {
    expect(splitVerticalOrientationRuns('2026')).toEqual([{ text: '2026', mode: 'sideways' }]);
  });

  it('returns nothing for empty text', () => {
    expect(splitVerticalOrientationRuns('')).toEqual([]);
  });

  it('preserves surrogate pairs as single code points', () => {
    const pieces = splitVerticalOrientationRuns('𠀋'); // CJK Ext-B ideograph (surrogate pair)
    expect(pieces).toHaveLength(1);
    expect(pieces[0].text).toBe('𠀋');
  });
});

// A minimal 2D-context spy recording the transform + text/box draw ops so we can
// assert the draw geometry without a real canvas.
type Op =
  | { op: 'save' }
  | { op: 'restore' }
  | { op: 'translate'; x: number; y: number }
  | { op: 'rotate'; a: number }
  | { op: 'scale'; sx: number; sy: number }
  | { op: 'fillText'; text: string; x: number; y: number; align: string; baseline: string }
  | { op: 'draw'; dx: number; dy: number; dw: number; dh: number };

// Optional metrics the mock returns from `measureText`, keyed by the metric it
// is asked to model. `fontBoundingBox*` are font-level (glyph-independent, used
// for the sideways em-box-centre shift); `actualBoundingBox*` are per-glyph tight
// ink bounds (used for the upright along-column ink centring). Values are read at
// the CURRENT textBaseline the helper sets before measuring, so the mock reports
// the same numbers regardless — the helper does the (asc−desc)/2 arithmetic.
interface MockMetrics {
  fontBoundingBoxAscent?: number;
  fontBoundingBoxDescent?: number;
  // Per-glyph tight ink extent under a `middle` textBaseline, by draw glyph.
  inkMiddle?: Record<string, { asc: number; desc: number }>;
}

function mockCtx(metrics: MockMetrics = {}): { ctx: any; ops: Op[] } {
  const ops: Op[] = [];
  const ctx: any = {
    textAlign: 'start',
    textBaseline: 'alphabetic',
    letterSpacing: '0px',
    save() {
      ops.push({ op: 'save' });
    },
    restore() {
      ops.push({ op: 'restore' });
    },
    translate(x: number, y: number) {
      ops.push({ op: 'translate', x, y });
    },
    rotate(a: number) {
      ops.push({ op: 'rotate', a });
    },
    scale(sx: number, sy: number) {
      ops.push({ op: 'scale', sx, sy });
    },
    measureText(s: string) {
      // Every code point is 10px wide.
      const m: Record<string, number> = { width: [...s].length * 10 };
      if (metrics.fontBoundingBoxAscent !== undefined) {
        m.fontBoundingBoxAscent = metrics.fontBoundingBoxAscent;
        m.fontBoundingBoxDescent = metrics.fontBoundingBoxDescent ?? 0;
      }
      const ink = metrics.inkMiddle?.[s];
      if (ink && this.textBaseline === 'middle') {
        m.actualBoundingBoxAscent = ink.asc;
        m.actualBoundingBoxDescent = ink.desc;
      }
      return m;
    },
    fillText(text: string, x: number, y: number) {
      ops.push({ op: 'fillText', text, x, y, align: this.textAlign, baseline: this.textBaseline });
    },
  };
  return { ctx, ops };
}

describe('drawVerticalRun (§17.6.20 — upright CJK counter-rotated, Latin sideways)', () => {
  it('counter-rotates every upright glyph −90° about its cell centre', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '富士', 100, 200, 12, 0);
    // Two upright glyphs → two save/rotate(−90°)/restore triples.
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    expect(rotates.every((r) => Math.abs(r.a - -Math.PI / 2) < 1e-9)).toBe(true);
    // First cell centre: x=100 + adv/2 (adv=10) = 105, baseline y=200.
    const firstTranslate = ops.find((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(firstTranslate).toEqual({ op: 'translate', x: 105, y: 200 });
    // Upright glyphs draw centred.
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.every((f) => f.align === 'center' && f.baseline === 'middle')).toBe(true);
  });

  it('draws a Latin glyph sideways (no rotation, alphabetic baseline, at the advance x)', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, 'A', 100, 200, 12, 0);
    expect(ops.some((o) => o.op === 'rotate')).toBe(false);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // Sideways: left at run x (advance 0), alphabetic baseline. The cross-axis y
    // is shifted down by the em-box centre so the glyph's ink centres on the same
    // column line the upright cells use. With no fontBoundingBox metrics the mock
    // falls back to 0.38·fontPx = 4.56, so y = 200 + 4.56.
    expect(fill?.text).toBe('A');
    expect(fill?.x).toBe(100);
    expect(fill?.baseline).toBe('alphabetic');
    expect(fill?.y).toBeCloseTo(200 + 0.38 * 12, 6);
  });

  it('centres a sideways glyph on the em-box centre from fontBoundingBox metrics (§17.6.20)', () => {
    // Symptom 1: mixed columns like "電話 03-…" must share one centreline. The
    // sideways glyph is drawn on its alphabetic baseline shifted down by the
    // font's em-box centre = (fontBoundingBoxAscent − fontBoundingBoxDescent)/2.
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 40, fontBoundingBoxDescent: 10 });
    drawVerticalRun(ctx, '0', 100, 200, 48, 0);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // emBoxCenter = (40 − 10)/2 = 15 → y = 200 + 15 = 215.
    expect(fill).toMatchObject({ text: '0', x: 100, baseline: 'alphabetic' });
    expect(fill?.y).toBeCloseTo(215, 6);
  });

  it('advances each glyph by measure + letterSpacing (measure == draw)', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, 'AB', 0, 0, 12, 4); // adv = 10 + 4 = 14 per glyph
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((f) => f.x)).toEqual([0, 14]);
  });

  it('rotates a Tr glyph WITHOUT a vertical form (ー) with the page — centred, NOT counter-rotated', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, 'ー', 100, 200, 12, 0);
    // ー (U+30FC) has no U+FE3x vertical form, so it keeps the rotate fallback:
    // the page rotation only (no −90° counter-rotation), centred on the column at
    // the cell centre (105, 200) with center/middle alignment.
    expect(ops.some((o) => o.op === 'rotate')).toBe(false);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill).toMatchObject({ text: 'ー', x: 105, y: 200, align: 'center', baseline: 'middle' });
  });

  it('substitutes a Tr bracket with its vertical form (（→︵, ）→︶) and draws it UPRIGHT', () => {
    // Symptom 2: a rotated fullwidth bracket lands its ink off-cell (unmeasurable
    // on a Canvas). Substituting the U+FE35/FE36 vertical form and drawing it
    // upright (counter-rotated) lets a per-glyph vertical-ink metric centre it.
    const { ctx, ops } = mockCtx({
      // Model ︵/︶ ink hugging opposite cell ends under a `middle` baseline.
      inkMiddle: { '︵': { asc: -9, desc: 21 }, '︶': { asc: 21, desc: -9 } },
    });
    drawVerticalRun(ctx, '（）', 0, 0, 12, 0);
    // Both are substituted to their vertical forms and COUNTER-ROTATED (upright).
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((f) => f.text)).toEqual(['︵', '︶']);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    expect(rotates.every((r) => Math.abs(r.a - -Math.PI / 2) < 1e-9)).toBe(true);
    // Along-column ink centring: fillText y = (asc − desc)/2. For ︵: (−9−21)/2 =
    // −15; for ︶: (21−(−9))/2 = 15. Drawn centred (x = 0, the cell centre).
    expect(fills.map((f) => f.align)).toEqual(['center', 'center']);
    expect(fills.map((f) => f.baseline)).toEqual(['middle', 'middle']);
    expect(fills.map((f) => f.x)).toEqual([0, 0]);
    expect(fills[0].y).toBeCloseTo(-15, 6);
    expect(fills[1].y).toBeCloseTo(15, 6);
  });

  it('keeps a Tr bracket at the cell centre when the Canvas reports no ink metrics', () => {
    // No actualBoundingBox* → along-column correction is 0, so the substituted
    // vertical form draws at the cell centre exactly (graceful degradation).
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '「」', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((f) => f.text)).toEqual(['﹁', '﹂']); // FE41 / FE42
    expect(fills.every((f) => f.x === 0 && f.y === 0)).toBe(true);
  });

  it('substitutes a Tu comma/period with its vertical presentation form (、→︑, 。→︒)', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '、。', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // Drawn glyphs are the vertical forms U+FE11 / U+FE12; both counter-rotated.
    expect(fills.map((f) => f.text)).toEqual(['︑', '︒']);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    // Substituted forms are pre-positioned by the font → drawn at the cell centre
    // with no upper-right nudge (local x offset 0). With no ink metrics the mock
    // reports none, so the along-column correction is 0 (y = 0) too.
    expect(fills.every((f) => f.x === 0 && f.y === 0)).toBe(true);
  });

  it('does NOT ink-centre a substituted Tu comma even when ink metrics are present', () => {
    // The comma/full stop vertical forms (︑ ︒) are DESIGNED with their ink in the
    // cell's upper-right corner (JIS X 4051). Unlike a Tr bracket, they must NOT be
    // pulled to the geometric cell centre by the along-column ink metric — doing so
    // drops them LOW (the "、。 sit too low" defect, #771). So even with ink metrics
    // reported, the punctuation draws at the em-box centre (y = 0), letting the
    // font's designed corner offset stand.
    const { ctx, ops } = mockCtx({
      // Model ︑ ink hugging the top of the cell (as a real vertical comma does).
      inkMiddle: { '︑': { asc: 21, desc: -9 } },
    });
    drawVerticalRun(ctx, '、', 0, 0, 12, 0);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill?.text).toBe('︑');
    // y stays 0 (em-box centred) — the ink metric is deliberately ignored here.
    expect(fill?.y).toBe(0);
    // Contrast: a Tr bracket WITH the same ink metric IS shifted (see the ︵/︶ test
    // above), proving the skip is punctuation-specific, not a blanket disable.
  });

  it('draws ！ / ？ UPRIGHT without substitution (they stand centred, not corner-hung)', () => {
    // ！ FF01 and ？ FF1F are vo=Tu but NOT substituted to FE15/FE16: those vertical
    // forms are corner-designed in many fonts and would push the mark off the column
    // centre (the sample-26 "！ shifted right" defect, #771). The original fullwidth
    // mark drawn upright is already the correct, centred vertical appearance.
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '！？', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // Drawn as the ORIGINAL code points (no FE15/FE16 substitution), counter-rotated
    // upright and centred on the cell (center/middle at x=0, y=0).
    expect(fills.map((f) => f.text)).toEqual(['！', '？']);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    expect(fills.every((f) => f.align === 'center' && f.baseline === 'middle')).toBe(true);
    expect(fills.every((f) => f.x === 0 && f.y === 0)).toBe(true);
  });
});

describe('drawTateChuYokoRun (§17.3.2.10 — horizontal-in-vertical / 縦中横)', () => {
  it('draws the whole run as ONE upright, centred fillText in the cell centre', () => {
    // Two full-width digits "２９" in a 12px cell (cellAdvance=12), no scale.
    const { ctx, ops } = mockCtx();
    drawTateChuYokoRun(ctx, '２９', 100, 200, 12, 12, 1, false);
    // Exactly one fillText — the whole run, not per glyph.
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills).toHaveLength(1);
    expect(fills[0].text).toBe('２９');
    // Centred (center/middle) at local origin.
    expect(fills[0]).toMatchObject({ align: 'center', baseline: 'middle', x: 0, y: 0 });
    // Counter-rotated −90° (upright, cancels the +90° page rotation).
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(1);
    expect(rotates[0].a).toBeCloseTo(-Math.PI / 2, 12);
    // Pivot = cell centre along the column: x = 100 + cellAdvance/2 = 106, y=200.
    const translate = ops.find((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translate).toEqual({ op: 'translate', x: 106, y: 200 });
  });

  it('applies w:w only to the cross-column (glyph width) axis — scale(charScale, 1)', () => {
    // §17.3.2.43 w:w=67 compresses the digits side-by-side ACROSS the column
    // (local x), NOT the along-column cell height (local y stays 1).
    const { ctx, ops } = mockCtx();
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 0.67, false);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale).toEqual({ op: 'scale', sx: 0.67, sy: 1 });
  });

  it('leaves the along-column axis at 1 when vertCompress content already fits one em', () => {
    // vertCompress on a single-line run whose natural height (asc+desc) equals the
    // em: no compression needed, so scale y stays 1 (the common 2-digit case).
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 3 }); // 12 = 1em
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 1, true);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale).toEqual({ op: 'scale', sx: 1, sy: 1 });
  });

  it('compresses the along-column axis when vertCompress content is taller than one em', () => {
    // A run whose natural upright height exceeds one em is squeezed to fit one
    // cell (§17.3.2.10): scale y = fontPx / height. Here height = 18 > 12 → 12/18.
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 14, fontBoundingBoxDescent: 4 }); // 18px
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 1, true);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale?.sx).toBe(1);
    expect(scale?.sy).toBeCloseTo(12 / 18, 12);
  });

  it('does not compress the height when vertCompress is off, even if content is tall', () => {
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 14, fontBoundingBoxDescent: 4 });
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 1, false);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale).toEqual({ op: 'scale', sx: 1, sy: 1 });
  });

  it('combines w:w (width) and vertCompress (height) independently', () => {
    // 3-digit-style case: w:w=0.5 across the column AND a tall run compressed to
    // one em along the column — the two axes are set independently.
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 18, fontBoundingBoxDescent: 6 }); // 24px
    drawTateChuYokoRun(ctx, '２９９', 0, 0, 12, 12, 0.5, true);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale?.sx).toBe(0.5);
    expect(scale?.sy).toBeCloseTo(12 / 24, 12);
  });
});

describe('drawUprightBox (§17.6.20 — keep images upright inside the rotated page)', () => {
  it('rotates −90° about the box centre and passes the swapped local box', () => {
    const { ctx, ops } = mockCtx();
    let called: number[] | null = null;
    drawUprightBox(ctx, 10, 20, 100, 40, (dx, dy, dw, dh) => {
      called = [dx, dy, dw, dh];
    });
    expect(ops).toContainEqual({ op: 'translate', x: 60, y: 40 }); // centre (10+50, 20+20)
    expect(ops).toContainEqual({ op: 'rotate', a: -Math.PI / 2 });
    // Local box: width↔height swap, centred on the pivot → (−h/2, −w/2, h, w).
    expect(called).toEqual([-20, -50, 40, 100]);
    // Balanced save/restore.
    expect(ops[0]).toEqual({ op: 'save' });
    expect(ops[ops.length - 1]).toEqual({ op: 'restore' });
  });
});

describe('physicalToLogicalAnchorBox (§17.6.20 + §20.4.3.x — physical anchor ↦ logical flow)', () => {
  it('projects a physical image box into the swapped logical frame (w↔h swap)', () => {
    // sample-26 ground truth (px at scale=1 = pt): physical page width 842pt,
    // image physical TL (444.3, 397.85), size 96.2 × 123.0. Word's physical
    // centroid (PDF-verified) is (492.4, 459.35).
    const cssW = 842;
    const box = physicalToLogicalAnchorBox(444.3, 397.85, 96.2, 123.0, cssW);
    // logical.x = physical.y ; logical.y = cssW − (physical.x + w) ; w↔h swap.
    expect(box.x).toBeCloseTo(397.85, 5);
    expect(box.y).toBeCloseTo(842 - (444.3 + 96.2), 5); // 301.5
    expect(box.w).toBeCloseTo(123.0, 5);
    expect(box.h).toBeCloseTo(96.2, 5);
  });

  it('round-trips: drawUprightBox on the logical box lands the image at the physical centroid', () => {
    // Feed the logical box through drawUprightBox and reconstruct the physical
    // rectangle by composing the page transform (translate(cssW,0)·rotate(+90))
    // with drawUprightBox's own (translate(centre)·rotate(−90)) — the net must be
    // the physical image box, upright.
    const cssW = 842;
    const px = 444.3;
    const py = 397.85;
    const w = 96.2;
    const h = 123.0;
    const box = physicalToLogicalAnchorBox(px, py, w, h, cssW);
    const { ctx, ops } = mockCtx();
    let local: number[] | null = null;
    drawUprightBox(ctx, box.x, box.y, box.w, box.h, (dx, dy, dw, dh) => {
      local = [dx, dy, dw, dh];
    });
    // The draw rect corners, transformed through page·drawUprightBox, must span
    // the physical image box.
    const translate = ops.find(
      (o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate',
    );
    // Compose: physical = P · translate(cx,cy) · rotate(−90) · localCorner.
    const cx = translate?.x ?? 0;
    const cy = translate?.y ?? 0;
    const P = (lx: number, ly: number): [number, number] => {
      // page transform: translate(cssW,0) then rotate(+90): (x,y) → (cssW−y, x)
      return [cssW - ly, lx];
    };
    const boxLocal = (dx: number, dy: number): [number, number] => {
      // drawUprightBox frame: translate(cx,cy)·rotate(−90): (x,y)→(cx+y, cy−x)
      const rx = cx + dy;
      const ry = cy - dx;
      return P(rx, ry);
    };
    const [dx, dy, dw, dh] = local as unknown as number[];
    const corners = [
      boxLocal(dx, dy),
      boxLocal(dx + dw, dy),
      boxLocal(dx, dy + dh),
      boxLocal(dx + dw, dy + dh),
    ];
    const xs = corners.map((c) => c[0]);
    const ys = corners.map((c) => c[1]);
    expect(Math.min(...xs)).toBeCloseTo(px, 4); // physical left
    expect(Math.min(...ys)).toBeCloseTo(py, 4); // physical top
    expect(Math.max(...xs) - Math.min(...xs)).toBeCloseTo(w, 4); // physical width
    expect(Math.max(...ys) - Math.min(...ys)).toBeCloseTo(h, 4); // physical height
    // Centroid matches Word / PDF ground truth.
    expect((Math.min(...xs) + Math.max(...xs)) / 2).toBeCloseTo(492.4, 3);
    expect((Math.min(...ys) + Math.max(...ys)) / 2).toBeCloseTo(459.35, 3);
  });
});

describe('verticalTextLayerPlacement (§17.6.20 — overlay span physical placement)', () => {
  it('maps a logical run top-left to the physical rotated placement', () => {
    // Logical run at (100, 200) on an 842px-wide physical page.
    const place = verticalTextLayerPlacement(100, 200, 842, true);
    expect(place).toEqual({ left: 842 - 200, top: 100, transform: 'rotate(90deg)' });
  });

  it('returns null on a horizontal page (span placed at logical x/y, no transform)', () => {
    expect(verticalTextLayerPlacement(100, 200, 842, false)).toBeNull();
  });
});
