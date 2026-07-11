import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import { drawEaVertRun } from './vertical-text.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §20.1.10.83 ST_TextVerticalType `eaVert` ("East Asian Vertical"):
// "some fonts are displayed as if rotated by 90 degrees while some fonts (mostly
// East Asian) are displayed vertical." The renderer rotates the whole text body
// +90° and lays it out horizontally in the rotated frame — so a plain fillText
// paints each glyph lying on its side (rotated with the page). Before this fix
// EVERY glyph, including CJK ideographs/kana (UAX#50 vo=U), was left sideways —
// the issue #790 defect. The fix drives each glyph through the core UAX#50
// classifier (`verticalOrientation` + the two vertical-form substitution maps):
//   • vo=U / vo=Tu  → UPRIGHT: counter-rotate −90° about the cell centre so the
//     glyph stands up (net rotation 0), substituting the U+FE1x vertical form for
//     the corner-hanging comma / full stop.
//   • vo=Tr with a U+FE3x vertical form (（）「」…) → SUBSTITUTE that form, drawn
//     UPRIGHT (UAX#50 §5 Tr = "substitute a vertical glyph, rotate only as
//     fallback").
//   • vo=Tr with NO vertical form (ー U+30FC, quotes) → ROTATE with the page
//     (net rotation +90°, the fallback).
//   • vo=R (Latin / digits) → stay SIDEWAYS (net rotation +90°), the conventional
//     "縦中横 not applied" appearance.
// `vert` / `vert270` are unchanged (all glyphs rotate — that IS their spec
// meaning), and the horizontal path is byte-identical (all of this is gated on
// the eaVert flow).

const FONT_PX = 20;
const SCALE = 1 / 12700; // emuToPx(emu, SCALE) = emu·SCALE; PT_TO_EMU=12700 ⇒ 1pt → 1px

// Representative glyphs, one per UAX#50 class (verified against the core table).
const U_CJK = '国'; // vo=U  → upright
const U_KANA = 'あ'; // vo=U  → upright
const R_LATIN = 'A'; // vo=R  → sideways
const TR_BRACKET = '（'; // vo=Tr → substitute U+FE35 ︵, upright
const TR_BRACKET_FE = String.fromCodePoint(0xfe35); // ︵
const TU_COMMA = '、'; // vo=Tu → substitute U+FE11, upright
const TU_COMMA_FE = String.fromCodePoint(0xfe11);
const TR_ROTATE = 'ー'; // vo=Tr, no vertical form → rotate 90°
// vo=Tr punctuation / white lenticular brackets with a U+FE1x form (issue #969).
const TR_VFORMS: Array<[string, string]> = [
  ['：', String.fromCodePoint(0xfe13)], // fullwidth colon → ︓
  ['；', String.fromCodePoint(0xfe14)], // fullwidth semicolon → ︔
  ['〖', String.fromCodePoint(0xfe17)], // left white lenticular → ︗
  ['〗', String.fromCodePoint(0xfe18)], // right white lenticular → ︘
];

interface DrawCall {
  text: string;
  x: number;
  y: number;
  /** Net canvas rotation in effect at draw time, normalised to (−π, π]. */
  rot: number;
  /** Accumulated translate-x in effect at draw time — the along-column cell
   *  centre for an upright glyph (translate(cx, …) precedes its fillText). */
  tx: number;
}

/** Normalise an angle to (−π, π] so 0 (upright) and π/2 (sideways) compare cleanly. */
function norm(a: number): number {
  return Math.atan2(Math.sin(a), Math.cos(a));
}

/** Recording 2D context that tracks the accumulated rotation through a
 *  save/restore stack, so each fillText records the NET rotation the glyph is
 *  painted under — 0 ≈ upright, +π/2 ≈ sideways/rotated. */
function mockCtx(): { ctx: CanvasRenderingContext2D; calls: DrawCall[] } {
  let font = `${FONT_PX}px serif`;
  let fillStyle = '';
  let letterSpacing = '0px';
  let direction: CanvasDirection = 'ltr';
  let textAlign: CanvasTextAlign = 'left';
  let textBaseline: CanvasTextBaseline = 'alphabetic';
  let rotation = 0;
  let tx = 0;
  const stack: { rotation: number; tx: number }[] = [];
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const calls: DrawCall[] = [];
  const metricsFor = (s: string): TextMetrics => {
    const p = px();
    // Full-width EA glyphs advance one em; ASCII ~half em.
    let w = 0;
    for (const ch of s) w += (ch.codePointAt(0) ?? 0) < 0x80 ? p * 0.5 : p;
    return {
      width: w,
      actualBoundingBoxAscent: p * 0.8,
      actualBoundingBoxDescent: p * 0.2,
      fontBoundingBoxAscent: p * 0.8,
      fontBoundingBoxDescent: p * 0.2,
    } as TextMetrics;
  };
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get textAlign() { return textAlign; }, set textAlign(v: CanvasTextAlign) { textAlign = v; },
    get textBaseline() { return textBaseline; }, set textBaseline(v: CanvasTextBaseline) { textBaseline = v; },
    measureText: (s: string) => metricsFor(s),
    fillText: (t: string, x: number, y: number) => calls.push({ text: t, x, y, rot: rotation, tx }),
    strokeText: (t: string, x: number, y: number) => calls.push({ text: t, x, y, rot: rotation, tx }),
    save: () => { stack.push({ rotation, tx }); },
    restore: () => { const s = stack.pop(); if (s) { rotation = s.rotation; tx = s.tx; } },
    translate: (x: number) => { tx += x; },
    rotate: (a: number) => { rotation += a; },
    scale: () => {},
    beginPath: () => {}, moveTo: () => {}, lineTo: () => {}, stroke: () => {},
    clip: () => {}, rect: () => {}, fillRect: () => {}, drawImage: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Serif',
  } as TextRunData;
}

function eaVertBody(text: string, alignment: Paragraph['alignment'] = 'l'): TextBody {
  const para: Paragraph = {
    alignment, marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], eaLnBrk: true, runs: [run(text)],
  } as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'none', vert: 'eaVert', autoFit: 'none',
  };
}

function renderEaVert(text: string, alignment: Paragraph['alignment'] = 'l', boxH = 400): DrawCall[] {
  const { ctx, calls } = mockCtx();
  // In eaVert the column length is the box HEIGHT; a tall box lets a `dist` line
  // stretch. bw is the cross-axis (column thickness).
  renderTextBody(ctx, eaVertBody(text, alignment), 0, 0, 200, boxH, SCALE);
  return calls;
}

const UPRIGHT = 0;
const SIDEWAYS = Math.PI / 2;

describe('pptx eaVert — UAX#50 per-glyph orientation (§20.1.10.83, issue #790)', () => {
  it('draws vo=U CJK ideographs UPRIGHT (net rotation 0), not sideways', () => {
    const calls = renderEaVert(U_CJK);
    const cjk = calls.filter((c) => c.text === U_CJK);
    expect(cjk.length, 'the ideograph is painted as its own glyph').toBe(1);
    expect(norm(cjk[0].rot), 'CJK stands upright (net rotation ≈ 0)').toBeCloseTo(UPRIGHT, 5);
  });

  it('draws vo=U kana UPRIGHT (net rotation 0)', () => {
    const calls = renderEaVert(U_KANA);
    const kana = calls.filter((c) => c.text === U_KANA);
    expect(kana.length).toBe(1);
    expect(norm(kana[0].rot)).toBeCloseTo(UPRIGHT, 5);
  });

  it('leaves vo=R Latin SIDEWAYS (net rotation +90°)', () => {
    const calls = renderEaVert(R_LATIN);
    const latin = calls.filter((c) => c.text === R_LATIN);
    expect(latin.length).toBe(1);
    expect(norm(latin[0].rot), 'Latin rotates with the page (sideways)').toBeCloseTo(SIDEWAYS, 5);
  });

  it('SUBSTITUTES a vo=Tr fullwidth bracket with its vertical form, drawn upright', () => {
    const calls = renderEaVert(TR_BRACKET);
    // No un-substituted fullwidth bracket is painted.
    expect(calls.some((c) => c.text === TR_BRACKET), 'original （ is not painted').toBe(false);
    const fe = calls.filter((c) => c.text === TR_BRACKET_FE);
    expect(fe.length, 'the U+FE35 vertical form ︵ is painted').toBe(1);
    expect(norm(fe[0].rot), 'substituted bracket is upright').toBeCloseTo(UPRIGHT, 5);
  });

  it('SUBSTITUTES a vo=Tu comma with its U+FE11 vertical form, drawn upright', () => {
    const calls = renderEaVert(TU_COMMA);
    expect(calls.some((c) => c.text === TU_COMMA), 'original 、 is not painted').toBe(false);
    const fe = calls.filter((c) => c.text === TU_COMMA_FE);
    expect(fe.length, 'the U+FE11 vertical form is painted').toBe(1);
    expect(norm(fe[0].rot)).toBeCloseTo(UPRIGHT, 5);
  });

  it.each(TR_VFORMS)(
    'SUBSTITUTES the vo=Tr colon/semicolon/lenticular %s with its U+FE1x vertical form, drawn upright (issue #969)',
    (orig, fe) => {
      const calls = renderEaVert(orig);
      expect(calls.some((c) => c.text === orig), `original ${orig} is not painted`).toBe(false);
      const sub = calls.filter((c) => c.text === fe);
      expect(sub.length, `the vertical form ${fe} is painted`).toBe(1);
      expect(norm(sub[0].rot), 'substituted form is upright').toBeCloseTo(UPRIGHT, 5);
    },
  );

  it('ROTATES a vo=Tr glyph with no vertical form (ー U+30FC) with the page (+90°)', () => {
    const calls = renderEaVert(TR_ROTATE);
    const mark = calls.filter((c) => c.text === TR_ROTATE);
    expect(mark.length, 'ー is painted as its own glyph').toBe(1);
    expect(norm(mark[0].rot), 'ー rotates 90° (the Tr fallback)').toBeCloseTo(SIDEWAYS, 5);
  });

  it('orients a mixed column: CJK upright, Latin sideways, bracket substituted, comma substituted, ー rotated', () => {
    const calls = renderEaVert(`${U_CJK}${R_LATIN}${TR_BRACKET}${TU_COMMA}${TR_ROTATE}`);
    const at = (t: string) => calls.find((c) => c.text === t);
    expect(norm(at(U_CJK)!.rot)).toBeCloseTo(UPRIGHT, 5);
    expect(norm(at(R_LATIN)!.rot)).toBeCloseTo(SIDEWAYS, 5);
    expect(at(TR_BRACKET_FE), 'bracket substituted').toBeTruthy();
    expect(norm(at(TR_BRACKET_FE)!.rot)).toBeCloseTo(UPRIGHT, 5);
    expect(at(TU_COMMA_FE), 'comma substituted').toBeTruthy();
    expect(norm(at(TU_COMMA_FE)!.rot)).toBeCloseTo(UPRIGHT, 5);
    expect(norm(at(TR_ROTATE)!.rot)).toBeCloseTo(SIDEWAYS, 5);
  });
});

// Justified / distributed vertical columns: `dist` (and `just` on a non-last
// pure-CJK line) opens a uniform pitch between glyphs to fill the column length.
// The eaVert draw must apply that pitch so the column SPREADS instead of bunching
// at the top (issue #790 codex review, finding 2 — the fully-distributed case).
describe('pptx eaVert — distributed columns spread the glyph pitch', () => {
  it('draws a `dist` pure-CJK column with a wider cell pitch than a left-aligned one', () => {
    const TEXT = 'あいうえお';
    const leftCol = renderEaVert(TEXT, 'l', 400).filter((c) => norm(c.rot) === 0 && [...TEXT].includes(c.text));
    const distCol = renderEaVert(TEXT, 'dist', 400).filter((c) => norm(c.rot) === 0 && [...TEXT].includes(c.text));
    expect(leftCol.length, 'all CJK glyphs upright').toBe(5);
    expect(distCol.length).toBe(5);
    const pitch = (col: DrawCall[]) => col[1].tx - col[0].tx; // along-column cell-centre step
    // Left-aligned: natural pitch ≈ one em (FONT_PX). Distributed: stretched to
    // fill the 400px column, so the pitch is materially larger.
    expect(pitch(leftCol)).toBeCloseTo(FONT_PX, 2);
    expect(pitch(distCol), 'distributed pitch fills the column').toBeGreaterThan(FONT_PX * 1.5);
  });
});

// Focused unit test of the draw helper in isolation (no page rotation installed,
// so the helper's OWN counter-rotation is what stands the glyph up). Here an
// upright glyph nets −90° (the helper's counter-rotation) and a sideways glyph
// nets 0 (the helper leaves it as the page would have drawn it).
describe('drawEaVertRun — per-glyph orientation helper', () => {
  function runHelper(text: string): DrawCall[] {
    const { ctx, calls } = mockCtx();
    drawEaVertRun(ctx, text, 0, 100, FONT_PX, 0, 'fill');
    return calls;
  }
  it('counter-rotates vo=U glyphs by −90° (upright in the +90° page frame)', () => {
    const calls = runHelper(U_CJK);
    expect(norm(calls[0].rot)).toBeCloseTo(-Math.PI / 2, 5);
    expect(calls[0].text).toBe(U_CJK);
  });
  it('does not counter-rotate vo=R Latin (stays with the page)', () => {
    const calls = runHelper(R_LATIN);
    expect(norm(calls[0].rot)).toBeCloseTo(0, 5);
    expect(calls[0].text).toBe(R_LATIN);
  });
  it('substitutes vo=Tr brackets and vo=Tu commas, counter-rotated upright', () => {
    expect(runHelper(TR_BRACKET)[0].text).toBe(TR_BRACKET_FE);
    expect(norm(runHelper(TR_BRACKET)[0].rot)).toBeCloseTo(-Math.PI / 2, 5);
    expect(runHelper(TU_COMMA)[0].text).toBe(TU_COMMA_FE);
    expect(norm(runHelper(TU_COMMA)[0].rot)).toBeCloseTo(-Math.PI / 2, 5);
  });
  it('leaves vo=Tr ー rotated with the page (no counter-rotation, no substitution)', () => {
    const calls = runHelper(TR_ROTATE);
    expect(calls[0].text).toBe(TR_ROTATE);
    expect(norm(calls[0].rot)).toBeCloseTo(0, 5);
  });
  it('advances each cell by measure + letterSpacingPx (the justification pitch)', () => {
    const { ctx: c0, calls: k0 } = mockCtx();
    drawEaVertRun(c0, '国国', 0, 100, FONT_PX, 0, 'fill');
    const { ctx: c1, calls: k1 } = mockCtx();
    drawEaVertRun(c1, '国国', 0, 100, FONT_PX, 8, 'fill');
    // Cell-centre step = advance = measure(国)=FONT_PX plus the per-glyph pitch.
    expect(k0[1].tx - k0[0].tx).toBeCloseTo(FONT_PX, 5);
    expect(k1[1].tx - k1[0].tx).toBeCloseTo(FONT_PX + 8, 5);
  });
});
