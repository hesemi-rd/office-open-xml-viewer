import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// Interop behaviour (NOT ECMA-376 — see isSectionBreakSpacerAt): an EMPTY
// paragraph that carries a section break (an inkless paragraph immediately
// followed by a `sectionBreak` element) has its spacing-BEFORE suppressed. Word
// and LibreOffice both render such a "section-break spacer" flush below the
// preceding paragraph (sample-13: the empty `mSectionBreak` between "Keywords"
// and "1. INTRODUCTION" carries before=22pt but neither editor applies it). A
// normal empty paragraph (not followed by a section break) keeps its before.

interface Call { text: string; y: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    strokeText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function para(text: string, spaceBefore = 0): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
        } as DocParagraph['runs'][number]]
      : [],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function docOf(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 600,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    // The final section (containing B) starts CONTINUOUS so the section break is
    // not a page break and B stays on page 0 (§17.6.22 — the break is governed by
    // the FOLLOWING section's start type).
    sectionStart: 'continuous',
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function baselineOf(body: BodyElement[], text: string): Promise<number> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(docOf(body), canvas, 0, { dpr: 1, width: 400 });
  const c = calls.find((k) => k.text === text);
  expect(c, `expected to paint ${text}`).toBeDefined();
  return (c as Call).y;
}

const SPACER_BEFORE = 20;

describe('section-break spacer suppresses spacing-before (Word/LibreOffice interop)', () => {
  it('an empty paragraph followed by a section break drops its 20pt before', async () => {
    // [A] [empty before=20] [sectionBreak continuous] [B]
    const spacerBody: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement,
      { type: 'sectionBreak', kind: 'continuous' } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    // Control: same, but the empty paragraph is NOT followed by a section break,
    // so it is a normal empty paragraph and keeps its 20pt before.
    const controlBody: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('', SPACER_BEFORE) as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];

    const aSpacer = await baselineOf(spacerBody, 'A');
    const bSpacer = await baselineOf(spacerBody, 'B');
    const aControl = await baselineOf(controlBody, 'A');
    const bControl = await baselineOf(controlBody, 'B');

    // 'A' is at the same place in both (nothing above it changed).
    expect(aSpacer).toBeCloseTo(aControl, 3);
    // The control applies the empty paragraph's 20pt before; the section-break
    // spacer suppresses it, so B sits exactly 20pt higher.
    expect(bControl - bSpacer).toBeCloseTo(SPACER_BEFORE, 1);
  });

  it('a NON-empty paragraph followed by a section break keeps its before (only empty spacers are suppressed)', async () => {
    // The section-ending paragraph here has text, so it is not an inkless spacer.
    const withText: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('X', SPACER_BEFORE) as unknown as BodyElement,
      { type: 'sectionBreak', kind: 'continuous' } as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const control: BodyElement[] = [
      para('A') as unknown as BodyElement,
      para('X', SPACER_BEFORE) as unknown as BodyElement,
      para('B') as unknown as BodyElement,
    ];
    const xWith = await baselineOf(withText, 'X');
    const xControl = await baselineOf(control, 'X');
    // The non-empty paragraph keeps its before regardless of the following break.
    expect(xWith).toBeCloseTo(xControl, 1);
  });
});
