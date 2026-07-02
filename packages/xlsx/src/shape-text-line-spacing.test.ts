import { describe, it, expect } from 'vitest';
import { PT_TO_PX, intendedSingleLinePx } from '@silurus/ooxml-core';
import { drawShapeText } from './renderer.js';
import type { ShapeParagraph, ShapeText, ShapeTextRun } from './types.js';

// ECMA-376 §21.1.2.2.5 <a:lnSpc> + §21.1.2.1.3 normAutofit lnSpcReduction for
// shape text bodies (drawShapeText). No real xlsx sample carries lnSpc or an
// applied autofit, so the feature is inert on the VRT corpus; these tests drive
// it directly with a mock CanvasRenderingContext2D, mirroring the (already
// verified) pptx model. 20 pt Calibri is used so intendedSingleLinePx returns 0
// (untabled face) and the natural single line is exactly the flat 1.2×em base.

interface FillTextCall {
  text: string;
  x: number;
  y: number;
}

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y });
    },
    drawImage() {},
    fillStyle: '#000' as string | CanvasGradient | CanvasPattern,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function textRun(text: string, size = 20, fontFace?: string): ShapeTextRun {
  return { type: 'text', text, bold: false, italic: false, size, fontFace };
}

function para(text: string, spaceLine?: ShapeParagraph['spaceLine'], fontFace?: string): ShapeParagraph {
  return { align: 'l', runs: [textRun(text, 20, fontFace)], spaceLine };
}

/** A paragraph whose single run declares only an East-Asian face (`<a:ea>`),
 *  latin left undefined — the common Japanese shape-text encoding. */
function paraEa(text: string, fontFaceEa: string): ShapeParagraph {
  return { align: 'l', runs: [{ type: 'text', text, bold: false, italic: false, size: 20, fontFaceEa }] };
}

/** A paragraph whose single run declares a tabled complex-script face (`<a:cs>`)
 *  but untabled latin/ea — the cs face must NOT floor the line box (it renders
 *  only complex-script glyphs). */
function paraCs(text: string, fontFaceCs: string): ShapeParagraph {
  return { align: 'l', runs: [{ type: 'text', text, bold: false, italic: false, size: 20, fontFaceCs }] };
}

/** Top-anchored, wrap:none, so a line's baseline is the cumulative sum of the
 *  heights of the lines above it. With two single-line paragraphs the A→B
 *  baseline gap equals the applied per-line height of the (identical) first line
 *  (textBaseline='middle': drawY = lineTop + height/2, so gap = h0/2 + h1/2). */
function gap(paragraphs: ShapeParagraph[], overrides: Partial<ShapeText> = {}): number {
  const { ctx, calls } = makeRecordingCtx();
  const txt: ShapeText = { anchor: 't', wrap: 'none', paragraphs, ...overrides };
  drawShapeText(ctx, txt, 400, 400, 1);
  const a = calls.find((c) => c.text === 'A');
  const b = calls.find((c) => c.text === 'B');
  expect(a).toBeDefined();
  expect(b).toBeDefined();
  return b!.y - a!.y;
}

describe('shape-text line spacing (§21.1.2.2.5 <a:lnSpc>) + normAutofit lnSpcReduction', () => {
  const cs = 1;
  const naturalSingle = 20 * PT_TO_PX * cs * 1.2; // 20 pt Calibri, floor is 0 → flat 1.2×em

  it('baseline: unspaced line height is the natural 1.2×em single line', () => {
    expect(gap([para('A'), para('B')])).toBeCloseTo(naturalSingle, 5);
  });

  it('spcPct 200% doubles the natural single-line height', () => {
    const spaced = gap([para('A', { type: 'pct', val: 200000 }), para('B', { type: 'pct', val: 200000 })]);
    expect(spaced).toBeCloseTo(naturalSingle * 2, 5);
    // And exactly 2× the unspaced reference.
    expect(spaced).toBeCloseTo(gap([para('A'), para('B')]) * 2, 5);
  });

  it('spcPts 40 makes each line an absolute 40 pt (cs-scaled) height', () => {
    const spaced = gap([para('A', { type: 'pts', val: 40 }), para('B', { type: 'pts', val: 40 })]);
    expect(spaced).toBeCloseTo(40 * PT_TO_PX * cs, 5);
  });

  it('normAutofit lnSpcReduction 0.2 scales each line to 80% of natural', () => {
    const reduced = gap([para('A'), para('B')], { autoFit: 'norm', lnSpcReduction: 0.2 });
    expect(reduced).toBeCloseTo(naturalSingle * 0.8, 5);
  });

  it('spAutoFit / noAutofit leave the natural line height unchanged (not applied)', () => {
    expect(gap([para('A'), para('B')], { autoFit: 'sp' })).toBeCloseTo(naturalSingle, 5);
    expect(gap([para('A'), para('B')], { autoFit: 'none' })).toBeCloseTo(naturalSingle, 5);
    // A stored fontScale is modeled but intentionally NOT applied to layout.
    expect(gap([para('A'), para('B')], { autoFit: 'norm', fontScale: 0.5 })).toBeCloseTo(naturalSingle, 5);
  });

  it('lnSpcReduction does NOT reduce absolute spcPts spacing (§21.1.2.1.3 note)', () => {
    // The spec: lnSpcReduction "applies only to paragraphs with percentage line
    // spacing." An absolute spcPts line height must be left as-is; only pct and
    // the implicit single (100 % percentage) are reduced.
    const pts = [para('A', { type: 'pts', val: 40 }), para('B', { type: 'pts', val: 40 })];
    const reduced = gap(pts, { autoFit: 'norm', lnSpcReduction: 0.2 });
    // Still the absolute 40 pt — NOT 40 × 0.8.
    expect(reduced).toBeCloseTo(40 * PT_TO_PX * cs, 5);
    // A pct paragraph in the same body IS reduced (control), proving the gate is
    // on the spacing type, not on the reduction being ignored entirely.
    const pct = [para('A', { type: 'pct', val: 100000 }), para('B', { type: 'pct', val: 100000 })];
    expect(gap(pct, { autoFit: 'norm', lnSpcReduction: 0.2 })).toBeCloseTo(naturalSingle * 0.8, 5);
  });
});

// Commit-1 companion: the natural single line is FLOORED by the authored font's
// design line (intendedSingleLinePx) before line spacing is applied. A tabled
// face (Meiryo, 1.5962×em) must measure to its taller design box; an untabled
// face (Calibri) stays on the flat 1.2×em. The mock canvas is metric-agnostic,
// so this exercises the pure design-metric floor keyed on the authored name.
describe('shape-text single-line height floor by font design metric (§21.1.2.1.1)', () => {
  const em = 20 * PT_TO_PX;

  it('a tabled face (Meiryo) floors the line to its design single line, above 1.2×em', () => {
    const meiryoFloor = intendedSingleLinePx('Meiryo', em);
    expect(meiryoFloor).toBeGreaterThan(em * 1.2); // sanity: the floor bites
    const h = gap([para('A', undefined, 'Meiryo'), para('B', undefined, 'Meiryo')]);
    expect(h).toBeCloseTo(meiryoFloor, 5);
  });

  it('an untabled face (Calibri) is left on the flat 1.2×em (floor returns 0)', () => {
    expect(intendedSingleLinePx('Calibri', em)).toBe(0);
    const h = gap([para('A', undefined, 'Calibri'), para('B', undefined, 'Calibri')]);
    expect(h).toBeCloseTo(em * 1.2, 5);
  });

  // ECMA-376 §21.1.2.3.1: a tabled face declared ONLY on `<a:ea>` (latin left
  // default) — the common Japanese shape-text encoding — must still floor the
  // line box by that face's design line. Before parsing `<a:ea>` the run's
  // fontFace stayed undefined and the floor never fired (PR #643 follow-up).
  it('a tabled EA face (Meiryo on <a:ea>, latin undefined) floors the line', () => {
    const meiryoFloor = intendedSingleLinePx('Meiryo', em);
    expect(meiryoFloor).toBeGreaterThan(em * 1.2); // sanity: the ea floor bites
    const h = gap([paraEa('A', 'Meiryo'), paraEa('B', 'Meiryo')]);
    expect(h).toBeCloseTo(meiryoFloor, 5);
  });

  it('a Calibri-latin run with no <a:ea> stays on the flat 1.2×em', () => {
    // Regression guard for the ea floor: an untabled latin face without any ea
    // must not grow (both floors return 0).
    const h = gap([para('A', undefined, 'Calibri'), para('B', undefined, 'Calibri')]);
    expect(h).toBeCloseTo(em * 1.2, 5);
  });

  // §21.1.2.3.1 font slots: the complex-script (`<a:cs>`) face renders ONLY
  // complex-script glyphs (Arabic/Hebrew/Thai). A tabled cs face on a run whose
  // glyphs are Latin/CJK must NOT floor the line box (that would over-grow it by
  // a face that renders none of the text — e.g. sample-25's Japanese run that
  // also declares Meiryo UI on cs). cs is parsed/modeled but excluded from the
  // line-box floor; correct cs handling is deferred to per-glyph layout.
  it('a tabled CS face (Meiryo on <a:cs>) does NOT floor the line box', () => {
    expect(intendedSingleLinePx('Meiryo', em)).toBeGreaterThan(em * 1.2); // Meiryo IS tabled
    const h = gap([paraCs('A', 'Meiryo'), paraCs('B', 'Meiryo')]);
    expect(h).toBeCloseTo(em * 1.2, 5); // stays flat — cs is not in the floor
  });
});
