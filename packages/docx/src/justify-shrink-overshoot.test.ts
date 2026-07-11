import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Word-observed `both`/`distribute` line-fit behavior (issue #698 PDF) — the
// Knuth-Plass space-shrink drawable-space tolerance (`SPACE_SHRINK_RATIO`) must
// NOT admit an extra word onto a line that the draw pass will justify, and must
// remain available on a line the draw pass treats as non-justified. §17.18.44
// classifies ST_Jc values but does not mandate this fit gate. A separate per-font
// bias applies exclusively on justified lines; this synthetic `serif` face has
// zero bias:
//
// - A line that WILL justify (a non-final, non-manual-break line of a
//   `both`/kashida paragraph, or ANY line of `distribute`/`thaiDistribute`)
//   redistributes its slack by EXPANDING inter-word spaces — it is never drawn
//   compressed below natural width, so a candidate word whose natural advance
//   overflows the column must wrap. Observed Word behaviour (issue #698, PDF
//   ground truth) still stands: in a narrow justified column Word breaks at
//   natural fit rather than pulling up one more word.
// - A line the draw pass does NOT justify — the paragraph's true last line and a
//   line ending at a manual `<w:br/>` (§17.3.3.1) under `both`/kashida, and every
//   line of a non-justified paragraph — is drawn with the shrink-fit compression
//   the budget promises (`shrinkFitCompression`).
// - The allowances are exclusive per line: justified lines receive only the
//   Canvas-vs-Word face bias; non-justified lines receive only drawable trailing-
//   space shrink. Demo p3/p6 space-collapse evidence shows that adding both
//   double-counts tolerance and admits words the paint pass cannot fit. Georgia
//   retains its calibrated bias on justified demo lines; generic `serif` and
//   Times remain at zero for the #698 natural-fit gate.

const FONT_PX = 12; // linear stub: each code point advances FONT_PX at scale 1

/** Linear recording canvas: measureText advances FONT_PX per code point (space
 *  included), so a token's trailing-space width is exactly one FONT_PX. This makes
 *  the fit arithmetic explicit — see TEXT4/TEXT5 below. */
function makeLinearCanvas(): HTMLCanvasElement {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => Number(/([\d.]+)px/.exec(font)?.[1] ?? FONT_PX);
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    fontKerning: 'auto',
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
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return canvas as unknown as HTMLCanvasElement;
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'serif', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocRun[], alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: FONT_PX, defaultFontFamily: 'serif',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function textPara(text: string, alignment: DocParagraph['alignment']): BodyElement {
  return para([{ type: 'text', ...textRun(text) } as DocRun], alignment);
}

function section(pageWidth: number): SectionProps {
  return {
    pageWidth, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    docGridCharSpace: undefined,
  } as SectionProps;
}

function doc(el: BodyElement, sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body: [el],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

/** Render one paragraph and return its visual lines, top-to-bottom: each line is
 *  the concatenated text of the onTextRun segments sharing a baseline y. */
async function renderLines(el: BodyElement, pageWidth: number): Promise<string[]> {
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(el, section(pageWidth)), makeLinearCanvas(), 0, {
    dpr: 1,
    width: pageWidth, // scale = 1 px/pt
    onTextRun: (r) => { if (r.text && r.text.trim()) runs.push(r); },
  });
  const byY = new Map<number, DocxTextRunInfo[]>();
  for (const r of runs) {
    const key = Math.round(r.y);
    let arr = byY.get(key);
    if (!arr) { arr = []; byY.set(key, arr); }
    arr.push(r);
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, rs]) => rs.slice().sort((p, q) => p.x - q.x).map((r) => r.text).join(''));
}

const tokens = (line: string): number => (line.match(/AAAA/g) ?? []).length;

// Fit arithmetic (linear stub): each "AAAA " token advances 5·12 = 60px, the
// bare word 48px, its trailing space 12px. Testing the 4th word on a line
// already holding three tokens: currentWidth = 180, wForFit = 48 ⇒ natural end
// 228. Column = 225 ⇒ overflow 3px. The line carries Σ trailing-space = 36px, so
// the shrink budget (0.25·36 = 9px) admits the word WHEN the tolerance applies.
const TEXT4 = 'AAAA AAAA AAAA AAAA';      // marginal word is the paragraph-FINAL word
const TEXT5 = 'AAAA AAAA AAAA AAAA AAAA'; // marginal word is followed by more content
const COLUMN = 225;

describe('§17.18.44 — per-font advance bias and drawable space shrink', () => {
  it('wraps the marginal word on a line that will justify', async () => {
    // Word PDF ground truth for #698's narrow justified column still shows a
    // natural-fit break (3 tokens). Generic `serif` has no per-font bias, and
    // this line will justify, so it gets no drawable-space budget either.
    const lines = await renderLines(textPara(TEXT5, 'both'), COLUMN);
    expect(lines.length).toBe(2);
    expect(tokens(lines[0])).toBe(3);
  });

  it('keeps the SAME marginal word on a non-justified line (drawable shrink retained)', async () => {
    // left/center lines are drawn at (or compressed toward) natural spacing, so
    // the 3px overflow is absorbed by drawable trailing-space compression. The
    // same path guards centred single-line titles such as sample-10 p1.
    for (const alignment of ['left', 'center'] as const) {
      const lines = await renderLines(textPara(TEXT5, alignment), COLUMN);
      expect(lines.length, alignment).toBe(2);
      expect(tokens(lines[0]), alignment).toBe(4); // budget admits the 4th; 5th wraps
    }
  });

  it("keeps the budget on a `both` paragraph's TRUE LAST line (paint draws it non-justified)", async () => {
    // The marginal word is the paragraph's final word: admitting it makes this
    // the paragraph's last line, which the draw pass does NOT justify
    // (applyJustify excludes endsLogicalLine for `both`) — it is drawn with the
    // shrink-fit compression the budget promises. Measure must therefore admit
    // within the same budget (measure==paint), keeping the old behaviour.
    const lines = await renderLines(textPara(TEXT4, 'both'), COLUMN);
    expect(lines.length).toBe(1);
    expect(tokens(lines[0])).toBe(4);
  });

  it('keeps the budget on a `both` line ending at a manual <w:br/> (§17.3.3.1)', async () => {
    // A manual break terminates the logical line, which `both` leaves
    // non-justified exactly like the paragraph's last line — same budget rule.
    const el = para(
      [
        { type: 'text', ...textRun(TEXT4) } as DocRun,
        { type: 'break', breakType: 'line' } as DocRun,
        { type: 'text', ...textRun('BBBB') } as DocRun,
      ],
      'both',
    );
    const lines = await renderLines(el, COLUMN);
    expect(lines.length).toBe(2); // "AAAA ×4" ‖ "BBBB" — the break-line keeps its 4th token
    expect(tokens(lines[0])).toBe(4);
    expect(lines[1]).toContain('BBBB');
  });

  it('wraps the marginal word on a `distribute` last line', async () => {
    // Word PDF ground truth for #698 still requires the natural-fit break here:
    // §17.18.44 `distribute` stretches the final line, yielding 3 tokens across
    // 2 lines. `distribute` stretches even the logical last line, so the
    // drawable-space budget is suppressed; generic `serif` contributes no bias.
    const lines = await renderLines(textPara(TEXT4, 'distribute'), COLUMN);
    expect(lines.length).toBe(2);
    expect(tokens(lines[0])).toBe(3);
  });

  it('does not force a wrap when the justified content genuinely fits at natural width', async () => {
    const lines = await renderLines(textPara(TEXT4, 'both'), 240);
    expect(lines.length).toBe(1);
  });

  it('scales the justified Georgia bias budget by authored w:w', async () => {
    const el = para(
      [{
        type: 'text',
        ...textRun('AAAA AAAA AAAA AAAA', { fontFamily: 'Georgia', charScale: 0.5 }),
      } as DocRun],
      'both',
    );

    // Natural end = 84px. Correct scaled bias = 0.882px (83 + 0.882 < 84)
    // so the third token wraps; the old unscaled 1.764px budget over-admitted it.
    const lines = await renderLines(el, 83);
    expect(tokens(lines[0])).toBe(2);
  });
});
