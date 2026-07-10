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
// ECMA-376 §17.3.2.14 `<w:fitText>` (Manual Run Width) inside an RTL paragraph
// (§17.3.1.6 `<w:bidi>`), resolved against UAX#9 reordering (rule L2).
//
// Issue #920 (follow-up from the PR #918 adversarial review): a fit region's
// residual width and its inter-character gaps were resolved in the LTR segment
// walk with no RTL fixture to verify the placement under the visual reorder.
//
// The reading-frame contract (same class as the docx #830 / pptx #913 tab-stop
// mirroring):
//   • The region occupies exactly `w:val`, anchored at the LEADING edge — the
//     left in an LTR paragraph, the RIGHT in an RTL paragraph.
//   • The inter-character gaps are distributed EVENLY in reading order; the
//     region's glyphs pack from the leading edge and the residual pad trails
//     AFTER the last glyph (to the left under RTL).
//   • The multi-run region's segments reorder visually (the logical-last run
//     draws on the visual LEFT), and each segment keeps the same per-gap pitch.
//
// The recording canvas uses a fixed-width mock glyph (`fontSize` px), so every
// x / letterSpacing / direction is exact and font-independent — the same
// technique the LTR fitText paint test and the #830 RTL tab test use.
// ─────────────────────────────────────────────────────────────────────────────

interface FillCall {
  text: string;
  x: number;
  letterSpacing: number;
  direction: string;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '12px serif';
  let letterSpacing = '0px';
  let direction: CanvasDirection = 'ltr';
  const fills: FillCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) { letterSpacing = value; },
    get direction() { return direction; },
    set direction(value: CanvasDirection) { direction = value; },
    fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, scale() {}, translate() {},
    fillText(text: string, x: number) {
      fills.push({ text, x, letterSpacing: Number.parseFloat(letterSpacing), direction });
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

function paragraph(runs: DocxTextRun[], bidi: boolean): BodyElement {
  const para: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((run) => ({ type: 'text', ...run })),
    defaultFontSize: 12, defaultFontFamily: 'serif', widowControl: false,
    ...(bidi ? { bidi: true } : {}),
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...para };
}

function section(): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

async function render(
  runs: DocxTextRun[],
  bidi: boolean,
): Promise<{ fills: FillCall[]; boxes: DocxTextRunInfo[] }> {
  const { canvas, fills } = makeRecordingCanvas();
  const boxes: DocxTextRunInfo[] = [];
  const model = {
    section: section(),
    body: [paragraph(runs, bidi)],
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

// Page is 600 px wide with no margins ⇒ scale 1, content [0, 600]. A 2400-twip
// region is 120 px. Each mock glyph is 12 px. All positions below are exact.
const TARGET = 120;
const RIGHT = 600;

describe('ECMA-376 §17.3.2.14 fitText in an RTL paragraph (issue #920)', () => {
  it('LTR control: a 3-run region packs from the left, byte-identical to #918', async () => {
    const fit = { fitTextVal: 2400, fitTextId: -1431456512 };
    const { fills, boxes } = await render(
      [textRun('ABCD', fit), textRun('E', fit), textRun('F', fit)],
      false,
    );
    // Six glyphs, five gaps: perGap = (120 − 72) / 5 = 9.6.
    expect(fills.map((f) => f.letterSpacing)).toEqual([9.6, 9.6, 9.6]);
    expect(fills.map((f) => f.direction)).toEqual(['ltr', 'ltr', 'ltr']);
    // Glyphs paint left→right in logical order; the region fills [0, 120].
    expect(fills.map((f) => f.x)).toEqual([0, 86.4, 108]);
    expect(boxes.map((b) => b.x)).toEqual([0, 86.4, 108]);
    expect(boxes.reduce((s, b) => s + b.w, 0)).toBeCloseTo(TARGET, 9);
  });

  it('RTL: a 2-run region mirrors — logical-last run on the visual LEFT, gaps preserved', async () => {
    // Same-id linked runs, both cs (§17.3.2.30 w:rtl) Arabic. Region = 120 px.
    const fit = { fitTextVal: 2400, fitTextId: 7, rtl: true };
    const { fills, boxes } = await render(
      [textRun('ابج', fit), textRun('دهو', fit)],
      true,
    );
    // Six glyphs, five gaps: perGap 9.6, same as the LTR control (reading-frame
    // distribution is direction-independent).
    expect(fills.map((f) => f.letterSpacing)).toEqual([9.6, 9.6]);
    expect(fills.map((f) => f.direction)).toEqual(['rtl', 'rtl']);
    // Visual order: the logical-LAST run ('دهو', the region end) draws on the
    // visual LEFT; the logical-first ('ابج') on the visual RIGHT.
    expect(fills.map((f) => f.text)).toEqual(['دهو', 'ابج']);
    // The region occupies exactly the leading (right) 120 px: [480, 600].
    // The region-end segment (leftmost) carries NO trailing gap: width
    // 36 + 2·9.6 = 55.2 ⇒ box [480, 535.2]. The leading segment carries its
    // cross-run boundary gap: 36 + 3·9.6 = 64.8 ⇒ box [535.2, 600].
    const sorted = [...boxes].sort((a, b) => a.x - b.x);
    expect(sorted.map((b) => Number(b.x.toFixed(4)))).toEqual([480, 535.2]);
    expect(sorted.map((b) => Number(b.w.toFixed(4)))).toEqual([55.2, 64.8]);
    const left = Math.min(...boxes.map((b) => b.x));
    const right = Math.max(...boxes.map((b) => b.x + b.w));
    expect(left).toBeCloseTo(RIGHT - TARGET, 6);
    expect(right).toBeCloseTo(RIGHT, 6);
  });

  it('RTL: a 3-run region mirrors with the residual folded into the even gaps', async () => {
    const fit = { fitTextVal: 2400, fitTextId: 9, rtl: true };
    const { fills, boxes } = await render(
      [textRun('ابجد', fit), textRun('ه', fit), textRun('و', fit)],
      true,
    );
    expect(fills.map((f) => f.letterSpacing)).toEqual([9.6, 9.6, 9.6]);
    // logical [ابجد, ه, و] ⇒ visual [و, ه, ابجد].
    expect(fills.map((f) => f.text)).toEqual(['و', 'ه', 'ابجد']);
    const left = Math.min(...boxes.map((b) => b.x));
    const right = Math.max(...boxes.map((b) => b.x + b.w));
    expect(left).toBeCloseTo(RIGHT - TARGET, 6);
    expect(right).toBeCloseTo(RIGHT, 6);
  });

  it('LTR: a single-char region leaves the residual as trailing pad on the RIGHT', async () => {
    // One glyph, no gaps: the whole (120 − 12) = 108 px residual is trailing pad.
    const { fills, boxes } = await render([textRun('A', { fitTextVal: 2400 })], false);
    expect(fills).toHaveLength(1);
    // The glyph sits at the leading (left) edge; the pad fills to the right.
    expect(fills[0].x).toBeCloseTo(0, 6);
    expect(boxes[0].x).toBeCloseTo(0, 6);
    expect(boxes[0].w).toBeCloseTo(TARGET, 6);
  });

  it('RTL: a single-char region puts the glyph at the leading (right) edge, pad to the LEFT', async () => {
    // The RTL mirror of the LTR case: the glyph must sit at the region's leading
    // (RIGHT) edge and the 108 px residual pad must trail to its LEFT — NOT sit
    // at the physical left with the pad bleeding rightward into the line.
    const { fills, boxes } = await render([textRun('ا', { fitTextVal: 2400, rtl: true })], true);
    expect(fills).toHaveLength(1);
    expect(fills[0].direction).toBe('rtl');
    // The region box is still the leading 120 px: [480, 600].
    expect(boxes[0].x).toBeCloseTo(RIGHT - TARGET, 6);
    expect(boxes[0].w).toBeCloseTo(TARGET, 6);
    // The glyph is drawn at the region's RIGHT edge: leftEdge (480) + pad (108) =
    // 588, so the 12 px glyph ends on the leading margin 600. Before the fix the
    // glyph painted at 480 (physical left) with the pad trailing rightward.
    expect(fills[0].x).toBeCloseTo(RIGHT - TARGET + (TARGET - 12), 6);
    expect(fills[0].x).toBeCloseTo(588, 6);
  });
});
