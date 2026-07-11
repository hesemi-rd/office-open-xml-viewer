import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types.js';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.1.6 `<w:bidi>` — inter-word space at a run/segment boundary in
// an RTL paragraph (issue #929).
//
// Root cause: the paint pass drew each segment's WHOLE string (including its
// TRAILING space) with a single `fillText`, relying on `ctx.direction='rtl'` to
// place that trailing space on the segment's physical LEFT (toward the next
// reading word). Chrome's Canvas does reorder the trailing whitespace, but
// skia-canvas — the server/VRT/MCP rendering backend — does NOT: it left-anchors
// the logical string and appends the trailing space on the physical RIGHT. In an
// RTL paragraph the reading-next word sits to the physical LEFT, so the trailing
// space lands on the wrong (outer) side and the two words render touching. For a
// two-word run (e.g. an Arabic name cell "أحمد علي") the single inter-word gap
// collapses to ~0.
//
// The backend-independent fix: do NOT depend on Canvas's per-string bidi for the
// EDGE whitespace. For an RTL-direction segment that ends in whitespace, draw the
// trailing-whitespace-TRIMMED glyphs anchored at `boxLeft + trailingSpaceWidth`
// so the space always occupies the segment box's LEFT — identical output in both
// Chrome and skia. The pen advance and the segment box (`onTextRun` / decorations)
// are unchanged.
//
// The recording canvas uses a fixed-width mock glyph (`fontSize` px), so every
// x / text is exact and font-independent — the same technique the fitText RTL
// paint test (fit-text-rtl.test.ts) uses.
// ─────────────────────────────────────────────────────────────────────────────

interface FillCall {
  text: string;
  /** DEVICE x of the fill (the raw argument mapped through any active
   *  translate/scale), so anchors inside a §17.3.2.43 `ctx.scale` frame are
   *  directly comparable with the unscaled segment boxes. */
  x: number;
  direction: string;
  /** In-frame letterSpacing (px) at the time of the fill. */
  letterSpacingPx: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '12px serif';
  let letterSpacing = '0px';
  let direction: CanvasDirection = 'ltr';
  // Minimal transform tracking (translate + horizontal scale, the renderer's
  // §17.3.2.43 w:w paint frame): fillText records tx + sx·x so device positions
  // are exact. save/restore keep a stack.
  let tx = 0;
  let sx = 1;
  const stack: { tx: number; sx: number }[] = [];
  const fills: FillCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) { letterSpacing = value; },
    get direction() { return direction; },
    set direction(value: CanvasDirection) { direction = value; },
    fontKerning: 'auto',
    // Honors letterSpacing (one pitch per code point, matching the renderer's
    // per-cp accounting) so a measurement mistakenly taken under a paint
    // letterSpacing differs from the natural width — the mock can catch a
    // w:w / w:spacing measure-vs-paint mismatch (#949 review, item 2).
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
      const ls = Number.parseFloat(letterSpacing) || 0;
      const cps = [...text].length;
      return {
        width: cps * px + cps * ls,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ tx, sx }); },
    restore() { const f = stack.pop(); if (f) { tx = f.tx; sx = f.sx; } },
    scale(fx: number) { sx *= fx; },
    translate(dx: number) { tx += dx * sx; },
    fillText(text: string, x: number) {
      fills.push({
        text,
        x: tx + sx * x,
        direction,
        letterSpacingPx: Number.parseFloat(letterSpacing) || 0,
      });
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: 'serif', isLink: false, background: null,
    vertAlign: null, ...extra, hyperlink: extra.hyperlink ?? null,
  };
}

function paragraph(runs: DocxTextRun[]): BodyElement {
  const para: DocParagraph = {
    alignment: 'right', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((run) => ({ type: 'text', ...run })),
    defaultFontSize: 12, defaultFontFamily: 'serif', widowControl: false,
    bidi: true,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...para };
}

function section(): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

async function render(runs: DocxTextRun[]): Promise<{ fills: FillCall[]; boxes: DocxTextRunInfo[] }> {
  const { canvas, fills } = makeRecordingCanvas();
  const boxes: DocxTextRunInfo[] = [];
  const model = {
    section: section(),
    body: [paragraph(runs)],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { serif: 'roman' },
  } as unknown as DocxDocumentModel;
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1,
    width: 600,
    onTextRun: (box) => boxes.push(box),
  });
  return { fills, boxes };
}

// Two Arabic words. Each mock glyph is 12 px; a space is 12 px. "ابج"=3 glyphs,
// "دهو"=3 glyphs. Page 600 px, no margins ⇒ scale 1.
const W1 = 'ابج'; // reading-first word (visually rightmost under RTL)
const W2 = 'دهو'; // reading-second word (visually leftmost)
const SPACE = 12;

/** The glyph fill call for a given (trimmed) word. */
function glyphFill(fills: FillCall[], word: string): FillCall | undefined {
  return fills.find((f) => f.text.trim() === word && f.text.trim() !== '');
}

describe('ECMA-376 §17.3.1.6 RTL inter-word space at a run boundary (issue #929)', () => {
  it('single run "W1 W2": the reading-first word keeps a full space before the reading-second word', async () => {
    const { fills, boxes } = await render([textRun(`${W1} ${W2}`, { rtl: true })]);

    // Both segments draw RTL. Visual order: W2 (leftmost), then W1 (rightmost).
    const w1Box = boxes.find((b) => b.text.trim() === W1)!; // "ابج " box (includes trailing space)
    const w2Box = boxes.find((b) => b.text.trim() === W2)!; // "دهو" box
    expect(w1Box).toBeTruthy();
    expect(w2Box).toBeTruthy();
    // Adjacent boxes: W2 on the left, W1 immediately to its right.
    expect(w1Box.x).toBeCloseTo(w2Box.x + w2Box.w, 6);

    // The reading-first word's GLYPHS must be drawn with NO trailing whitespace,
    // anchored so its trailing space occupies the box's LEFT (toward W2).
    const w1Fill = glyphFill(fills, W1)!;
    expect(w1Fill).toBeTruthy();
    // (a) glyphs carry no trailing space (the space is positioned explicitly).
    expect(/\s$/u.test(w1Fill.text)).toBe(false);
    // (b) glyph anchor sits one space-width to the RIGHT of the box left edge.
    expect(w1Fill.x).toBeCloseTo(w1Box.x + SPACE, 6);

    // (c) The decisive visual invariant, backend-independent: the empty gap
    //     between W2's right edge and W1's first glyph equals exactly one space —
    //     the words do NOT touch. Before the fix W1's glyphs started at w1Box.x
    //     (flush against W2) and the space sat on the far (outer) right.
    const w2GlyphRight = w2Box.x + w2Box.w; // "دهو" has no trailing space
    expect(w1Fill.x - w2GlyphRight).toBeCloseTo(SPACE, 6);
  });

  it('two adjacent runs "W1 " + "W2": same full inter-word space (run boundary)', async () => {
    const { fills, boxes } = await render([
      textRun(`${W1} `, { rtl: true }),
      textRun(W2, { rtl: true }),
    ]);
    const w1Box = boxes.find((b) => b.text.trim() === W1)!;
    const w2Box = boxes.find((b) => b.text.trim() === W2)!;
    const w1Fill = glyphFill(fills, W1)!;
    expect(/\s$/u.test(w1Fill.text)).toBe(false);
    expect(w1Fill.x).toBeCloseTo(w1Box.x + SPACE, 6);
    expect(w1Fill.x - (w2Box.x + w2Box.w)).toBeCloseTo(SPACE, 6);
  });

  it('LTR control is unchanged: an LTR-direction trailing-space word still draws its full string flush-left', async () => {
    // Latin words in an LTR paragraph — the fast path must be byte-identical.
    const { canvas, fills } = makeRecordingCanvas();
    const boxes: DocxTextRunInfo[] = [];
    const para: DocParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [{ type: 'text', ...textRun('ABC DEF') }],
      defaultFontSize: 12, defaultFontFamily: 'serif', widowControl: false,
    } as unknown as DocParagraph;
    const model = {
      section: section(),
      body: [{ type: 'paragraph', ...para } as BodyElement],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: { serif: 'roman' },
    } as unknown as DocxDocumentModel;
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 600, onTextRun: (b) => boxes.push(b) });
    // "ABC " draws its full string (with trailing space) at the box's left edge.
    const abc = fills.find((f) => f.text.startsWith('ABC'))!;
    expect(abc.text).toBe('ABC ');
    expect(abc.direction).toBe('ltr');
    const abcBox = boxes.find((b) => b.text.trim() === 'ABC')!;
    expect(abc.x).toBeCloseTo(abcBox.x, 6);
  });

  it('w:w + w:spacing (§17.3.2.43 + §17.3.2.35): the anchor shift equals the AUTHORITY whitespace advance — the fixed pitch is NOT scaled by w:w', async () => {
    // charScale 0.5, charSpacing 2 pt (scale 1 ⇒ 2 px per code point).
    // Advance authority (segAdvanceWidth): natural × 0.5 + cp × 2.
    //   "ابج " → 48·0.5 + 4·2 = 32   |   "ابج" → 36·0.5 + 3·2 = 24  |  "دهو" → 24.
    // Whitespace advance = 32 − 24 = 8 px (= 12·0.5 glyph + 2 pitch). A
    // measurement mistakenly taken under `letterSpacing = spacing` and THEN
    // scaled yields shift 32 − (36+3·2)·0.5 = 11 — the mock's letterSpacing-aware
    // measureText makes that mistake visible.
    const { fills, boxes } = await render([
      textRun(`${W1} ${W2}`, { rtl: true, charScale: 0.5, charSpacing: 2 }),
    ]);
    const w1Box = boxes.find((b) => b.text.trim() === W1)!;
    const w2Box = boxes.find((b) => b.text.trim() === W2)!;
    expect(w1Box.w).toBeCloseTo(32, 6);
    expect(w2Box.w).toBeCloseTo(24, 6);
    // Boxes adjacent, W2 visually left of W1.
    expect(w1Box.x).toBeCloseTo(w2Box.x + w2Box.w, 6);

    const w1Fill = glyphFill(fills, W1)!;
    expect(/\s$/u.test(w1Fill.text)).toBe(false);
    // (a) anchor shift = the authority whitespace advance (8), not the
    //     letterSpacing-inflated 11.
    expect(w1Fill.x - w1Box.x).toBeCloseTo(8, 6);
    // (b) measure==paint: the trimmed glyphs' painted advance inside the ×0.5
    //     frame (letterSpacing = spacing/scale = 4 ⇒ 0.5·(36 + 3·4) = 24) lands
    //     the right edge exactly on the box edge.
    expect(w1Fill.letterSpacingPx).toBeCloseTo(4, 6);
    const paintedAdvance = 0.5 * (3 * SPACE + 3 * w1Fill.letterSpacingPx);
    expect(w1Fill.x + paintedAdvance).toBeCloseTo(w1Box.x + w1Box.w, 6);
    // (c) the whitespace-less W2 keeps its box-left anchor.
    const w2Fill = glyphFill(fills, W2)!;
    expect(w2Fill.x).toBeCloseTo(w2Box.x, 6);
  });

  it('fitText × RTL (§17.3.2.14, review of #949): a trailing space inside the fit region also falls to the glyphs’ LEFT', async () => {
    // Two same-id linked rtl runs forming one 120 px region (2400 twips at
    // scale 1): "ابج " + "دهو" = 7 cps, natural 84 px ⇒ 6 gaps, perGap = 6.
    //   "ابج " (non-end, 4 cps): 48 + 4·6 = 72   |   "دهو" (end, 3 cps): 36 + 2·6 = 48.
    // Whitespace advance (authority Δ) = 72 − (36 + 3·6) = 18 px (glyph 12 +
    // per-gap share 6).
    const fit = { fitTextVal: 2400, fitTextId: 31, rtl: true };
    const { fills, boxes } = await render([textRun(`${W1} `, fit), textRun(W2, fit)]);
    const w1Box = boxes.find((b) => b.text.trim() === W1)!;
    const w2Box = boxes.find((b) => b.text.trim() === W2)!;
    expect(w1Box.w).toBeCloseTo(72, 6);
    expect(w2Box.w).toBeCloseTo(48, 6);
    // Visual order: the logical-last run (دهو, the region end) on the LEFT.
    expect(w1Box.x).toBeCloseTo(w2Box.x + w2Box.w, 6);

    // The region-start segment's glyphs draw TRIMMED, shifted right by the
    // whitespace advance so the space (and its gap share) sits to their left;
    // the per-gap pitch is preserved. Before this fix the fitText branch drew
    // the UNTRIMMED string at the box left edge (skia: space stranded on the
    // outer right, gap to دهو collapsed).
    const w1Fill = glyphFill(fills, W1)!;
    expect(/\s$/u.test(w1Fill.text)).toBe(false);
    expect(w1Fill.letterSpacingPx).toBeCloseTo(6, 6);
    expect(w1Fill.x - w1Box.x).toBeCloseTo(18, 6);
    // measure==paint: trimmed advance 36 + 3·6 = 54 ⇒ right edge on the box edge.
    expect(w1Fill.x + (3 * SPACE + 3 * 6)).toBeCloseTo(w1Box.x + w1Box.w, 6);
    // The region-end segment is unaffected (no trailing whitespace).
    const w2Fill = glyphFill(fills, W2)!;
    expect(w2Fill.x).toBeCloseTo(w2Box.x, 6);
  });
});
