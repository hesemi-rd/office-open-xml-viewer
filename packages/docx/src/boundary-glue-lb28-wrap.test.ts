import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Regression for the "segment boundary = unconditional line-break opportunity"
// bug. `buildSegments` emits one layout segment per (run × word × script slice),
// so two ADJACENT runs with NO whitespace between them become two segments and
// the wrap loop was free to break between them — even where UAX#14 grants no
// break opportunity. The canonical case: a `<` (U+003C, Line_Break class AL)
// authored in its OWN run, immediately followed by an Arabic-letter run (also
// AL). UAX#14 LB28 (AL × AL) forbids a break there, so Word wraps at the PREVIOUS
// space and keeps `<` with the word that follows it; we orphaned `<` at the end
// of the previous line.
//
// The fix marks the following segment `joinPrev` (UAX#14 LB28, via the shared
// core `isUax14NoBreakPair`) so the existing atomic-group pre-flush keeps the two
// together and breaks at the previous legal opportunity (the space) instead.

const FONT_PX = 20; // uniform glyph advance in the stub (scale = 1)

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

function runsPara(texts: string[], pageWidth: number): { body: BodyElement; sec: SectionProps } {
  const p: DocParagraph = {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: texts.map((t) => ({ type: 'text', ...textRun(t) }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  };
  const sec: SectionProps = {
    pageWidth, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  };
  return { body: { type: 'paragraph', ...p } as BodyElement, sec };
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(texts: string[], pageWidth: number) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const { body, sec } = runsPara(texts, pageWidth);
  await renderDocumentToCanvas(doc([body], sec), canvas, 0, { dpr: 1, width: pageWidth });
  return fillTextCalls;
}

function linesByY(calls: { text: string; x: number; y: number }[]) {
  const byY = new Map<number, { text: string; x: number }[]>();
  for (const c of calls) {
    const k = Math.round(c.y);
    (byY.get(k) ?? byY.set(k, []).get(k)!).push({ text: c.text, x: c.x });
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, glyphs]) => glyphs.slice().sort((p, q) => p.x - q.x));
}

/** The line (array of glyph fragments) whose concatenated text contains `needle`. */
function lineContaining(lines: { text: string }[][], needle: string): string | undefined {
  const l = lines.find((g) => g.map((x) => x.text).join('').includes(needle));
  return l?.map((x) => x.text).join('');
}

describe('UAX#14 no-break pairs at a run boundary (LB28 + LB14/LB23/LB25/LB30)', () => {
  it('keeps "<" with the Latin word that immediately follows it (own runs, no space)', async () => {
    // 7 cells/line (140px / 20px). Runs: "wwww " (5) | "<" (1) | "xxxx" (4).
    //   "wwww " fills 5 cells; "<" fits (6); "<xxxx" (5) does NOT fit after it.
    // Pre-fix: "<" placed on line 1, "xxxx" flushed to line 2 → "<" orphaned.
    // Post-fix: "xxxx" is glued to "<" (LB28); the pair pre-flushes together, so
    //   line 1 = "wwww" and line 2 = "<xxxx".
    const calls = await render(['wwww ', '<', 'xxxx'], 140);
    expect(calls.length).toBeGreaterThan(0);
    const lines = linesByY(calls);

    const bracketLine = lineContaining(lines, '<');
    expect(bracketLine).toBeDefined();
    // The word that follows "<" must be on the SAME line as "<".
    expect(bracketLine).toContain('xxxx');
    // No line may BEGIN with the orphaned "<" alone while its word sits below it,
    // i.e. "<" and "xxxx" are never on different lines.
    const wordLine = lineContaining(lines, 'xxxx');
    expect(wordLine).toBe(bracketLine);
  });

  it('keeps "<" with the Arabic word that immediately follows it (report case)', async () => {
    // 5 cells/line (100px). Runs: "نص " (3) | "<" (1) | "شيء" (3).
    //   "نص " fills 3 cells; "<" fits (4); "<شيء" (4) fits alone but not after "نص ".
    // Pre-fix: "<" ends line 1, Arabic wraps to line 2 (LB28 boundary broken).
    // Post-fix: "<" + Arabic wrap together to line 2; line 1 = "نص".
    const calls = await render(['نص ', '<', 'شيء'], 100);
    expect(calls.length).toBeGreaterThan(0);
    const lines = linesByY(calls);

    const bracketLine = lineContaining(lines, '<');
    expect(bracketLine).toBeDefined();
    // "<" and the Arabic word share one line (order is bidi-reordered at paint,
    // but they must land in the same line box).
    expect(bracketLine).toContain('شيء');
  });

  it('keeps letters with the digits that immediately follow them (LB23 AL × NU)', async () => {
    // 7 cells/line (140px / 20px). Runs: "wwww " (5) | "Ab" (2) | "12" (2).
    //   "wwww Ab" fills 7 cells; "12" does not fit after it. LB23 forbids the
    //   AL × NU seam, so "Ab12" wraps down as one unit.
    const calls = await render(['wwww ', 'Ab', '12'], 140);
    const lines = linesByY(calls);
    const digitLine = lineContaining(lines, '12');
    expect(digitLine).toBeDefined();
    expect(digitLine).toContain('Ab12');
  });

  it('keeps a currency prefix with the digits that follow it (LB25 PR × NU)', async () => {
    // Runs: "wwww " (5) | "$" (1) | "100" (3): "wwww $" = 6 fits, "100" does not.
    // LB25 (PR × NU) forbids the seam, so "$100" moves down together.
    const calls = await render(['wwww ', '$', '100'], 140);
    const lines = linesByY(calls);
    const digitLine = lineContaining(lines, '100');
    expect(digitLine).toBeDefined();
    expect(digitLine).toContain('$100');
  });

  it('never orphans an opening bracket, whatever follows it (LB14 OP × NU)', async () => {
    // Runs: "wwww " (5) | "(" (1) | "2026" (4): "wwww (" = 6 fits, "2026" does
    // not. LB14 (OP SP* ×) forbids any break after "(", so "(2026" wraps as one.
    const calls = await render(['wwww ', '(', '2026'], 140);
    const lines = linesByY(calls);
    const digitLine = lineContaining(lines, '2026');
    expect(digitLine).toBeDefined();
    expect(digitLine).toContain('(2026');
  });

  it('keeps a word with the parenthetical that follows it (LB30 AL × OP)', async () => {
    // Runs: "www " (4) | "no" (2) | "(s)" (3): "www no" = 6 fits, "(s)" does not
    // (9 > 7). LB30 (AL × OP, non-East-Asian) forbids the n|( seam, so "no(s)"
    // moves down as one unit.
    const calls = await render(['www ', 'no', '(s)'], 140);
    const lines = linesByY(calls);
    const parenLine = lineContaining(lines, '(s)');
    expect(parenLine).toBeDefined();
    expect(parenLine).toContain('no(s)');
  });

  it('still breaks at a real whitespace boundary (does not over-glue)', async () => {
    // Runs: "wwww " (5) | "yyyy " (5) | "zzzz" (4), 7 cells/line.
    // Every boundary here follows a space, so each is a legal wrap opportunity —
    // the fix must NOT glue them. "wwww" / "yyyy" / "zzzz" land on separate lines.
    const calls = await render(['wwww ', 'yyyy ', 'zzzz'], 140);
    expect(calls.length).toBeGreaterThan(0);
    const lines = linesByY(calls);

    const yLine = lineContaining(lines, 'yyyy');
    const zLine = lineContaining(lines, 'zzzz');
    expect(yLine).toBeDefined();
    expect(zLine).toBeDefined();
    // The space between "yyyy " and "zzzz" is a legal break, so they wrap apart.
    expect(zLine).not.toBe(yLine);
    expect(yLine).not.toContain('zzzz');
  });
});
