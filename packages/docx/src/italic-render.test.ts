import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.2.16 `w:i` / §17.3.2.17 `w:iCs` — the run italic axis must
// survive into the paint-time font string (issue #928 regression guard).
//
// The renderer resolves the axis in buildSegments (line-layout.ts): non-complex-
// script content (Latin/Cyrillic/CJK — see core isComplexScriptCodePoint) takes
// `w:i` directly; content forced onto the complex-script axis (`w:rtl`
// §17.3.2.30 or `w:cs` §17.3.2.7) takes `w:iCs`. Inert CS METADATA on the run —
// §17.3.2.39 `szCs`, §17.3.2.20 `w:lang w:bidi` (both commonly stamped by Word's
// docDefaults) — must NOT reroute non-CS text onto the CS axis and shadow `w:i`;
// that shadowing was the suspected (disproven) mechanism of #928.
//
// These tests pin the paint font through renderDocumentToCanvas + onTextRun,
// whose `font` payload is the exact ctx.font the glyph draw used (buildFont
// output: "<style> <weight> <size>px <family…>").
//
// ADJUDICATED (issue #937): a force-CS run carrying `w:i`/`w:b` with `w:iCs`/
// `w:bCs` ABSENT paints UPRIGHT / non-bold. §17.3.2.17 `w:iCs` and §17.3.2.3
// `w:bCs` are INDEPENDENT toggles — the non-CS `w:i`/`w:b` value governs only
// the Latin axis and does NOT inherit onto the complex-script axis. Word's
// ground truth (fixture sample-41 cs-italic-toggle): Case A (`w:cs`+`w:i`, no
// `w:iCs`, Latin) and Case C (`w:rtl`+`w:i`, no `w:iCs`, Arabic) both render
// upright — identical to Case B (`w:iCs=0` explicit OFF) — while Case D (plain
// `w:i`, Latin axis) renders italic. The bold side is the same: sample-7's
// `w:rtl`+`w:cs`+`w:b` (no `w:bCs`) Arabic headings render at regular weight in
// Word's PDF, not bold. The renderer therefore resolves the CS axis as
// `italicCs ?? false` / `boldCs ?? false`, NOT `?? base.italic`/`?? base.bold`.
// ─────────────────────────────────────────────────────────────────────────────

const FONT_PX = 20;

/** Minimal 2D-context stub: stores whatever `ctx.font` string the renderer
 *  assigns (buildFont output) so onTextRun reports it verbatim. */
function makeRecordingCanvas(): HTMLCanvasElement {
  let font = `${FONT_PX}px serif`;
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
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
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    scale() {}, translate() {}, rotate() {}, transform() {}, setTransform() {}, resetTransform() {},
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return canvas as unknown as HTMLCanvasElement;
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'Times New Roman', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  } as DocxTextRun;
}

function para(runs: DocxTextRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r })),
    defaultFontSize: FONT_PX, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function doc(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 600, pageHeight: 600, marginTop: 0, marginRight: 0, marginBottom: 0,
    marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
    evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

/** Render one page and return the paint font (ctx.font at glyph draw) per text. */
async function paintFonts(body: BodyElement[]): Promise<Map<string, string>> {
  const fonts = new Map<string, string>();
  await renderDocumentToCanvas(doc(body), makeRecordingCanvas(), 0, {
    dpr: 1, width: 600,
    onTextRun: (r) => { if (!fonts.has(r.text)) fonts.set(r.text, r.font); },
  });
  return fonts;
}

describe('run italic axis at paint time (§17.3.2.16 w:i / §17.3.2.17 w:iCs)', () => {
  it('plain w:i run (no CS metadata) paints italic', async () => {
    const fonts = await paintFonts([para([textRun('курсив', { italic: true })])]);
    expect(fonts.get('курсив')).toMatch(/^italic /);
  });

  it('adjacent w:i and w:b runs keep independent style/weight', async () => {
    const fonts = await paintFonts([para([
      textRun('обычный'),
      textRun('полужирный', { bold: true }),
      textRun('курсивный', { italic: true }),
    ])]);
    expect(fonts.get('обычный')).toMatch(/^normal 400 /);
    expect(fonts.get('полужирный')).toMatch(/^normal 700 /);
    expect(fonts.get('курсивный')).toMatch(/^italic 400 /);
  });

  it('inert CS metadata (szCs + langBidi) does not shadow w:i on non-CS text', async () => {
    // The #928 fixture shape: a Cyrillic run whose docDefaults stamp szCs and
    // w:lang w:bidi. Neither is a §17.3.2.7/§17.3.2.30 CS trigger, so the run
    // stays on the non-CS axis and w:i must reach the paint font.
    const fonts = await paintFonts([para([
      textRun('курсив', { italic: true, fontSizeCs: FONT_PX, langBidi: 'ar-sa' }),
    ])]);
    expect(fonts.get('курсив')).toMatch(/^italic /);
  });

  it('force-CS run with explicit w:iCs=false paints upright (independent toggle)', async () => {
    // §17.3.2.7 <w:cs/> routes ALL of the run onto the CS axis, where an
    // explicit <w:iCs w:val="false"/> wins over <w:i/> — the non-CS toggle
    // governs disjoint content and must not leak through.
    const fonts = await paintFonts([para([
      textRun('курсив', { italic: true, cs: true, italicCs: false, fontSizeCs: FONT_PX }),
    ])]);
    expect(fonts.get('курсив')).toMatch(/^normal 400 /);
  });

  it('force-CS run (w:cs) with w:i and absent w:iCs paints upright (#937 Case A)', async () => {
    // §17.3.2.7 <w:cs/> routes the run onto the CS axis; §17.3.2.17 `w:iCs` is
    // an INDEPENDENT toggle, so an ABSENT `w:iCs` defaults OFF and `w:i` (Latin
    // axis) does not leak through. Word ground truth: sample-41 Case A renders
    // upright, identical to the explicit-OFF Case B above. Adjudicated in #937.
    const fonts = await paintFonts([para([
      textRun('курсив', { italic: true, cs: true, fontSizeCs: FONT_PX }),
    ])]);
    expect(fonts.get('курсив')).toMatch(/^normal 400 /);
  });

  it('w:rtl run with w:i and absent w:iCs paints upright (#937 Case C)', async () => {
    // The rtl (§17.3.2.30) sibling of Case A: same absent-`w:iCs` fallback.
    // Word ground truth: sample-41 Case C (Arabic) renders upright, not oblique.
    // Previously pinned to italic (`italicCs ?? italic`); flipped by the #937
    // adjudication to the independent-toggle reading (`italicCs ?? false`).
    const fonts = await paintFonts([para([
      textRun('نص', { italic: true, rtl: true, fontSizeCs: FONT_PX, langBidi: 'ar-sa' }),
    ])]);
    expect(fonts.get('نص')).toMatch(/^normal /);
  });

  it('force-CS run with w:b and absent w:bCs paints non-bold (#937 bold symmetry)', async () => {
    // §17.3.2.3 `w:bCs` mirrors §17.3.2.17 `w:iCs`: an INDEPENDENT toggle whose
    // absence defaults OFF, so `w:b` (Latin axis) does not force CS bold. Word
    // ground truth: sample-7's w:rtl+w:cs+w:b (no w:bCs) Arabic headings render
    // at regular weight, not bold. `csBold = boldCs ?? false`. Adjudicated #937.
    const fonts = await paintFonts([para([
      textRun('نص', { bold: true, rtl: true, fontSizeCs: FONT_PX, langBidi: 'ar-sa' }),
    ])]);
    expect(fonts.get('نص')).toMatch(/^normal 400 /);
  });
});
