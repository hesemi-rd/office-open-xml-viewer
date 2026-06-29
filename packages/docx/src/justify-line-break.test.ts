import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.18.44 (ST_Jc `both`) + §17.3.3.1 (`<w:br>` text-wrapping break):
// a line terminated by a MANUAL line break is the end of a logical line and is
// LEFT-aligned (not stretched) in a justified paragraph — exactly like the
// paragraph's final line. Word does this; sample-16 is one giant `both`-justified
// paragraph whose items are separated by `<w:br/>`, and each item's last line
// (e.g. "…インデントなし。") renders left-aligned in Word but was stretched
// (sparse) by us because the renderer only un-justified the paragraph's TRUE last
// line (`li === lines.length - 1`), not lines ending at a break.

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

/** A justified paragraph: a short run, a MANUAL line break, then a long run.
 *  The first line (the short run) ends at the break. For `both` it must NOT
 *  justify (logical line end → left-aligned); for `distribute` it STILL gets
 *  filled to the margin (§17.18.44, see the distribute test below). */
function breakPara(alignment: DocParagraph['alignment'] = 'both'): BodyElement {
  const p: DocParagraph = {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [
      { type: 'text', ...textRun('ああああ') } as DocRun,
      { type: 'break', breakType: 'line' } as DocRun,
      { type: 'text', ...textRun('いいいいいいいいいいいいいいいい') } as DocRun,
    ],
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(): SectionProps {
  return {
    pageWidth: 210, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    // Active char grid so every full-width glyph is drawn individually → one
    // fillText per glyph, letting us read the realised inter-glyph pitch.
    // Cell = 20 + (-2048/4096) = 19.5px; availW 210 → 10 cells fit.
    docGridType: 'linesAndChars', docGridLinePitch: 20, docGridCharSpace: -2048,
  } as SectionProps;
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

describe('justified paragraph — a line ended by a manual <w:br/> is left-aligned (§17.18.44 + §17.3.3.1)', () => {
  it('does not stretch the break-terminated first line', async () => {
    const calls = await render([breakPara()], section());
    expect(calls.length).toBeGreaterThan(0);

    // Group draws by baseline y; the first (top) line is the break-terminated one.
    const byY = new Map<number, { text: string; x: number }[]>();
    for (const c of calls) {
      const key = Math.round(c.y);
      (byY.get(key) ?? byY.set(key, []).get(key)!).push(c);
    }
    const firstY = Math.min(...byY.keys());
    const line = byY.get(firstY)!.slice().sort((p, q) => p.x - q.x);

    // A LEFT-aligned (un-justified) line has NO internal distribute gaps, so the
    // pure-EA grid segment "ああああ" is painted as ONE contiguous fillText (the
    // contextual-shaping path), NOT four isolated per-glyph draws.
    expect(line.length).toBe(1);
    expect(line[0].text).toBe('ああああ');

    // Left-aligned: it starts at the margin and the segment box ends at its
    // NATURAL grid width (4 cells ≈ 4 × FONT_PX), NOT stretched toward the 210px
    // right margin. Cell = 20 + (-2048/4096) = 19.5px ⇒ 4 cells = 78px.
    expect(line[0].x).toBeLessThan(FONT_PX); // starts at the left margin
    const naturalRight = 4 * FONT_PX; // ≈ box edge upper bound
    expect(line[0].x + 4 * (FONT_PX - 2048 / 4096)).toBeLessThan(naturalRight + FONT_PX);
  });

  // ECMA-376 §17.18.44 (ST_Jc): `distribute` justifies EVERY line — inter-word
  // AND inter-character — including the paragraph's final line. So unlike `both`
  // (which leaves a logical line end left-aligned), a `distribute` line that ends
  // at a manual `<w:br/>` is STILL stretched to the margin. The renderer encodes
  // this as the `stretchLastLine = (alignment === 'distribute')` carve-out:
  //   applyJustify = isJustified && (!endsLogicalLine || stretchLastLine)
  // This test guards that carve-out: with the SAME content as the `both` case
  // above, the break-terminated first line must spread toward the right margin.
  it('still stretches the break-terminated first line under distribute', async () => {
    const calls = await render([breakPara('distribute')], section());
    expect(calls.length).toBeGreaterThan(0);

    const byY = new Map<number, { text: string; x: number }[]>();
    for (const c of calls) {
      const key = Math.round(c.y);
      (byY.get(key) ?? byY.set(key, []).get(key)!).push(c);
    }
    const firstY = Math.min(...byY.keys());
    const line = byY.get(firstY)!.slice().sort((p, q) => p.x - q.x);

    // Same first line "ああああ" (4 glyphs) ended by the break.
    expect(line.length).toBe(4);
    expect(line[0].x).toBeLessThan(FONT_PX); // still starts at the left margin

    // Stretched: the last glyph is pushed FAR past its natural end (4 × FONT_PX
    // = 80px) toward the ~210px right margin. Observed last-glyph x is ~190px
    // (the 4th of 4 glyphs spread across the ~210px line, cell pitch ~63px).
    // Threshold 120 sits well above the natural 80px and well below the observed
    // 190px, so it robustly distinguishes "stretched" from "natural width".
    const naturalRight = 4 * FONT_PX; // 80px
    expect(line[3].x).toBeGreaterThan(120);
    expect(line[3].x).toBeGreaterThan(naturalRight);
    // Inter-glyph advances are spread WIDE (well beyond the natural pitch),
    // confirming inter-character distribution on the break-terminated line.
    for (let i = 1; i < line.length; i++) {
      expect(line[i].x - line[i - 1].x).toBeGreaterThan(FONT_PX * 1.5);
    }
  });
});
