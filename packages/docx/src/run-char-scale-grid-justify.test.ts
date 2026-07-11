import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.3.2.43 `<w:w>` (charScale) must be applied at PAINT time even
// when a docGrid charSpace (§17.6.5) or distributed justify (§17.18.44) also
// governs the run — the two prior arms folded charScale into the measured box
// (segAdvanceWidth) but drew glyphs at full width, so the ink overran the box
// (issue #816: probe box 55px vs ink 105px). This test pins measure==paint for
// both combinations by reconstructing the painted ink extent under the run's
// horizontal ctx.scale transform and comparing it to the reported run box.

const FONT_PX = 20; // full-width EA glyph advance in the stub (scale 1)

interface FillCall {
  text: string;
  x: number; // the raw x passed to fillText (LOCAL to any active translate/scale)
  y: number;
  letterSpacing: number; // parsed px
  scaleX: number; // active horizontal scale at draw time
  translateX: number; // active x translate at draw time
}

/** Recording 2D context that tracks a simple x-only transform stack so the run's
 *  w:w ctx.scale/translate is visible, and models per-glyph EA advance as
 *  len·FONT_PX (no contextual collapse — pure metrics, sufficient for the box
 *  invariant since the grid/justify arms measure BEFORE setting letterSpacing). */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: FillCall[] = [];
  let scaleX = 1;
  let translateX = 0;
  const stack: { scaleX: number; translateX: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ scaleX, translateX }); },
    restore() { const s = stack.pop(); if (s) { scaleX = s.scaleX; translateX = s.translateX; } },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    // translate then scale (the renderer does translate(x,0) then scale(sx,1)),
    // so the applied local→screen map is screen = translateX + scaleX·local.
    scale(sx: number) { scaleX *= sx; },
    translate(tx: number) { translateX += tx * scaleX; },
    fillText(text: string, x: number, y: number) {
      fills.push({ text, x, y, letterSpacing: parseFloat(letterSpacing), scaleX, translateX });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocxTextRun[], alignment: DocParagraph['alignment'] = 'left'): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics', widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 400, marginTop: 0, marginRight: 0, marginBottom: 0,
    marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
    evenAndOddHeaders: false, docGridCharSpace: undefined, ...overrides,
  } as SectionProps;
}

function charGrid(charSpace: number): Partial<SectionProps> {
  return { docGridType: 'linesAndChars', docGridLinePitch: 20, docGridCharSpace: charSpace };
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(
  body: BodyElement[],
  sec: SectionProps,
): Promise<{ runs: DocxTextRunInfo[]; fills: FillCall[] }> {
  const { canvas, fills } = makeRecordingCanvas();
  const info: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1, width: sec.pageWidth, onTextRun: (r) => info.push(r),
  });
  return { runs: info, fills };
}

/** The painted advance extent (screen px) of a run's glyph draws: from the
 *  leftmost drawn glyph origin to the rightmost glyph's post-advance pen, mapping
 *  every LOCAL coordinate through the active translate/scale. Canvas `letterSpacing`
 *  inserts spacing after EVERY glyph (including the last — the pen advances past
 *  it), so the per-piece local advance is n·(FONT_PX + letterSpacing); the run
 *  scale then maps it to screen. This models the pen the NEXT segment abuts, i.e.
 *  the reserved box the measure pass produced (measure==paint). */
function paintedInkExtent(fills: FillCall[], expectedText: string): { left: number; right: number } {
  // Pieces that belong to this run: same substring content (a grid/justify run
  // may be drawn as several contiguous pieces). Match by inclusion in the text.
  const cps = [...expectedText];
  const pieces = fills.filter((f) => f.text.length > 0 && [...f.text].every((ch) => cps.includes(ch)));
  expect(pieces.length, `at least one glyph draw for ${JSON.stringify(expectedText)}`).toBeGreaterThan(0);
  let left = Infinity;
  let right = -Infinity;
  for (const p of pieces) {
    const n = [...p.text].length;
    const localAdvance = n * FONT_PX + n * p.letterSpacing;
    const screenLeft = p.translateX + p.scaleX * p.x;
    const screenRight = p.translateX + p.scaleX * (p.x + localAdvance);
    left = Math.min(left, screenLeft);
    right = Math.max(right, screenRight);
  }
  return { left, right };
}

describe('WD4 #816 — w:w charScale reaches paint under an active docGrid / justify (measure==paint)', () => {
  // Reviewer's probe (grid arm): a pure-EA run with w:w=0.5 in a linesAndChars
  // grid. The measured box scales natural width by 0.5 and adds the cell delta;
  // the paint must draw the glyphs at 0.5× so the ink fills exactly that box.
  it('grid arm: charScale=0.5 pure-EA run paints ink == reported box (not full width)', async () => {
    const charSpace = 2048; // +0.5pt per cell at scale 1
    const deltaPx = charSpace / 4096;
    const scale = 0.5;
    const text = 'あいうえお'; // 5 pure-EA glyphs, one segment
    const n = [...text].length;

    const { runs, fills } = await render(
      [para([textRun(text, { charScale: scale })])],
      section(charGrid(charSpace)),
    );

    const seg = runs.find((r) => r.text === text);
    expect(seg, 'run box reported').toBeDefined();
    // MEASURE: natural×scale + n cells of grid delta.
    const box = n * FONT_PX * scale + n * deltaPx;
    expect(seg!.w).toBeCloseTo(box, 6);

    // PAINT: ink extent under the run scale equals the reported box.
    const ink = paintedInkExtent(fills, text);
    const inkWidth = ink.right - ink.left;
    expect(ink.left).toBeCloseTo(seg!.x, 4);
    expect(inkWidth, 'ink fills the box (not overrun at full width)').toBeCloseTo(seg!.w, 4);
    // Regression guard: at least one draw is under the horizontal scale.
    expect(fills.some((f) => Math.abs(f.scaleX - scale) < 1e-9)).toBe(true);
  });

  // Justify arm: a pure-EA run with w:w=0.5 that is also distributed-justified.
  // distributeLineSlack opens a gap at every inter-CJK boundary on a wrapped
  // (non-last) line; the box = natural×0.5 + slack, and the paint must scale the
  // glyphs while leaving the justify pitch un-stretched.
  it('justify arm: charScale=0.5 distributed run paints ink == reported box', async () => {
    const scale = 0.5;
    // A long pure-EA run (NO docGrid ⇒ the justify arm, not the grid arm) that
    // wraps on a page sized so the first line is justified (both) with POSITIVE
    // slack ⇒ gaps open at every inter-CJK boundary (fully distributed). At W=205
    // the scaled first line holds 20 glyphs (200 px) with ~5 px slack, so the
    // fully-distributed fast path fires (uniform letterSpacing = distPerGap). The
    // latent bug: that arm drew glyphs at FULL width (scaleX 1) while the box was
    // measured at natural×0.5 + slack ⇒ ink overran the box.
    const text = 'あいうえおかきくけこさしすせそたちつてとなにぬねの';
    const { runs, fills } = await render(
      [para([textRun(text, { charScale: scale })], 'both')],
      section({ pageWidth: 205 }),
    );

    // The first justified line's run — drawn with a non-zero justify pitch
    // (letterSpacing) AND under the run's horizontal scale (post-fix). Before the
    // fix the justify-pitched draw was at scaleX 1 (the ink overrun).
    const firstY = Math.min(...fills.map((f) => f.y));
    const firstLine = fills.filter((f) => f.y === firstY);
    const pitched = firstLine.find((f) => f.letterSpacing > 0);
    expect(pitched, 'the first line is fully-distributed (justify pitch present)').toBeDefined();
    // Regression guard: the justify-pitched draw is under the horizontal scale.
    expect(pitched!.scaleX, 'justify draw honours w:w scale').toBeCloseTo(scale, 9);

    // Match the reported run box for the drawn text.
    const seg = runs.find(
      (r) => pitched!.text.startsWith(r.text) || r.text === pitched!.text,
    );
    expect(seg, 'run box for the justified line').toBeDefined();

    // The fully-distributed line is drawn as ONE fillText carrying the whole run
    // under `ctx.scale(scale, 1)` with letterSpacing = distPerGap/scale. Unlike a
    // grid cell, the justify slack excludes the FINAL glyph's trailing gap (the
    // box edge is the last glyph's advance, not one gap past it), so the painted
    // segment span is n·FONT_PX·scale + (n−1)·distPerGap. That must equal the
    // reported box (measure==paint); before the fix the glyphs painted at FULL
    // width (n·FONT_PX + (n−1)·distPerGap) — a ~2× overrun.
    expect(pitched!.text, 'whole line drawn in one piece').toBe(seg!.text);
    const n = [...pitched!.text].length;
    const distPerGapScaled = pitched!.scaleX * pitched!.letterSpacing; // distPerGap
    const paintedSpan = n * FONT_PX * pitched!.scaleX + (n - 1) * distPerGapScaled;
    const screenLeft = pitched!.translateX + pitched!.scaleX * pitched!.x;
    expect(screenLeft).toBeCloseTo(seg!.x, 3);
    expect(paintedSpan, 'painted span fills the justified box (not full width)').toBeCloseTo(seg!.w, 3);
  });

  // Sanity: the plain-scale arm (no grid, no justify) is unchanged — the existing
  // run-char-metrics-render coverage — but re-pinned here through the same
  // ink-extent lens so a future refactor of the three arms stays consistent.
  it('plain arm: charScale=0.5 with no grid/justify still paints ink == box', async () => {
    const scale = 0.5;
    const text = 'WORD';
    const { runs, fills } = await render([para([textRun(text, { charScale: scale })])], section());
    const seg = runs.find((r) => r.text === text);
    expect(seg!.w).toBeCloseTo([...text].length * FONT_PX * scale, 6);
    const ink = paintedInkExtent(fills, text);
    expect(ink.right - ink.left).toBeCloseTo(seg!.w, 4);
  });
});
