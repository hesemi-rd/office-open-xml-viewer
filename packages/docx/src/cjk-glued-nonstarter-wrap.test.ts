import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Regression: a CJK run followed by a glued non-starter ("。" / "、" etc., in its
// OWN run with no intervening space) must still CJK-split to fill the line. The
// non-starter carries `joinPrev` (UAX#14 LB13, keeps "。" off a line head); the
// line-breaker's glued-group pre-flush — meant for ATOMIC Latin/small-caps groups
// like "system" + "," — wrongly treated {CJK-run, 。} as atomic too, flushing the
// WHOLE group to the next line instead of splitting the (breakable) CJK run. The
// prior line then carried far fewer characters and a justified paragraph stretched
// it wide (sample-9: line 2 held 27 of 39 chars, spacing ~46% too wide).
//
// The fix skips the pre-flush when the lead segment is CJK-breakable, so the CJK
// run splits normally and "。" stays with its tail (kinsoku still keeps it off the
// next line's head). This guards the docx side; pptx/xlsx do not share this path.

const FONT_PX = 20; // glyph advance per CJK char in the stub (scale = 1)

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number }[];
} {
  let font = `${FONT_PX}px serif`;
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
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
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) { fillTextCalls.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

/** A `both`-justified paragraph whose runs are: a CJK run that fills the line
 *  start, a CJK run that must split to fill the rest, and a glued non-starter
 *  "。" in its own run (→ joinPrev). */
function gluedPara(texts: string[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'both',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: texts.map((t) => ({ type: 'text', ...textRun(t) }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 400, // contentWidth 200 → exactly 10 CJK cells/line
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(body: BodyElement[], sec: SectionProps) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, { dpr: 1, width: sec.pageWidth });
  return fillTextCalls;
}

describe('CJK run + glued non-starter — fill the line instead of pushing the whole group down', () => {
  it('splits the CJK run onto the first line; does not orphan its space', async () => {
    // availW = 200 px = 10 cells. run A "あ"×6 (120px) fills the start; run B
    // "い"×6 must split (4 onto line 1 → full 10, 2 to line 2); run C "。" is a
    // glued non-starter (its own run → joinPrev). Pre-fix: the glued-group
    // pre-flush moved {B,。} whole to line 2, leaving line 1 with only A's 6 cells.
    const calls = await render(
      [gluedPara(['ああああああ', 'いいいいいい', '。'])],
      section(),
    );
    expect(calls.length).toBeGreaterThan(0);

    // First (top) line.
    const byY = new Map<number, { text: string; x: number }[]>();
    for (const c of calls) {
      const k = Math.round(c.y);
      (byY.get(k) ?? byY.set(k, []).get(k)!).push({ text: c.text, x: c.x });
    }
    const firstY = Math.min(...byY.keys());
    const line1 = byY.get(firstY)!.slice().sort((p, q) => p.x - q.x);
    // Count CHARACTERS, not fillText calls: a justified CJK line is drawn in
    // pieces (one call may cover several glyphs), so the call count is not the
    // cell count.
    const line1Chars = [...line1.map((g) => g.text).join('')];

    // The breakable CJK run B must split to fill line 1 (10 cells), not be pushed
    // down whole behind its glued "。". Pre-fix line 1 held 6 cells (only run A).
    expect(line1Chars.length).toBe(10);
    expect(line1Chars).toContain('い'); // run B reached line 1
    expect(line1Chars).not.toContain('。'); // the non-starter stays with its tail
  });

  it('keeps the glued "。" off the next line head when the CJK run exactly fills the line', async () => {
    // The contract the fix leans on: with the pre-flush skipped, a CJK run that
    // EXACTLY fills the line (10 cells) followed by a standalone "。" must not
    // orphan "。" at the next line's head — §17.3.1.16 kinsoku retracts the run's
    // last cell so "。" follows a CJK char ("あ。"), never leading a line.
    const calls = await render([gluedPara(['ああああああああああ', '。'])], section());
    expect(calls.length).toBeGreaterThan(0);

    const byY = new Map<number, { text: string; x: number }[]>();
    for (const c of calls) {
      const k = Math.round(c.y);
      (byY.get(k) ?? byY.set(k, []).get(k)!).push({ text: c.text, x: c.x });
    }
    // Every line: its leading (min-x) glyph must not be the non-starter "。".
    for (const [, glyphs] of byY) {
      const head = glyphs.slice().sort((p, q) => p.x - q.x)[0];
      expect(head.text).not.toBe('。');
    }
    // And "。" is actually present (it was rendered, not dropped).
    expect(calls.map((c) => c.text).join('')).toContain('。');
  });
});
