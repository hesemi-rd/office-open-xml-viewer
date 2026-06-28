import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, HeaderFooter, SectionProps } from './types';

// ECMA-376 §17.6.11 (pgMar/@top) — the SYMMETRIC twin of the footer rule in
// footer-reserve.test.ts. The main-document text TOP is placed at the GREATER of the
// top margin and the header's extent ("The value of top / The extent of the header
// text"), so a header taller than its top-margin allowance (marginTop − headerDistance)
// rises into the content area and the body must start BELOW it: Word never lays main
// text over a header ("the main text extent ends at the bottom of the header region").
// The paginator reserves that overflow at the TOP of every page's content area and the
// body's first line is pushed down by the same amount. A NEGATIVE top margin is the
// spec's explicit exception — the main text is then "measured from the top of the page
// extent regardless of the header ... and therefore shall overlap the header text", so
// nothing is reserved. These tests pin those rules with a synthetic doc whose header is
// far taller than its top margin (the header-side mirror of sample-13's masthead).

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

function para(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
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

// pageHeight 600, margins 10, headerDistance 4 → top-margin allowance for the header
// is marginTop − headerDistance = 6pt. A header taller than 6pt overflows the top.
function docWithHeader(
  body: BodyElement[],
  header: HeaderFooter | null,
  opts: { marginTop?: number } = {},
): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 600,
    marginTop: opts.marginTop ?? 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage',
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: header, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

async function renderPage0(doc: DocxDocumentModel): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('header reserve — content never overlaps a tall header (ECMA-376 §17.6.11)', () => {
  // A few single-line body paragraphs (page 0 holds them comfortably) — enough to read
  // the topmost body line. The reserve, not a long body, is what places that line.
  const body = (): BodyElement[] =>
    Array.from({ length: 6 }, () => para('BODY') as unknown as BodyElement);
  // A header far taller than the 6pt allowance (6 lines ≈ ~70pt), so its bottom edge
  // sinks well below the top margin and into the body's region.
  const tallHeader: HeaderFooter = {
    body: Array.from({ length: 6 }, () => para('HDR') as unknown as BodyElement),
  };

  it('starts the body below a tall header so no line is painted into the header band', async () => {
    const withHeader = await renderPage0(docWithHeader(body(), tallHeader));
    const noHeader = await renderPage0(docWithHeader(body(), null));

    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const headerY = (calls: Call[]) => calls.filter((c) => c.text === 'HDR').map((c) => c.y);

    const minBodyTall = Math.min(...bodyY(withHeader));
    const maxHeader = Math.max(...headerY(withHeader));
    const minBodyNone = Math.min(...bodyY(noHeader));

    // Sanity: the tall header is actually painted on page 0.
    expect(headerY(withHeader).length).toBeGreaterThan(0);

    // INVARIANT: with the tall header reserved, the topmost body line on page 0 sits
    // BELOW the bottom-most header line — body never overlaps the header.
    expect(minBodyTall).toBeGreaterThan(maxHeader);

    // NON-TRIVIALITY: the SAME body, with no header to reserve against, starts at the
    // top margin — ABOVE where the tall header's bottom sits — proving the invariant
    // above is not satisfied trivially and that the reservation, not a short header, is
    // what clears the header band.
    expect(minBodyNone).toBeLessThan(maxHeader);
  });

  it('does not reserve a header when the top margin is negative (§17.6.11 exception)', async () => {
    // §17.6.11: a negative top margin measures the main text from the page top
    // REGARDLESS of the header, so the text overlaps the header and nothing is
    // reserved. Hence the same body starts at exactly the same y WITH the tall header
    // as WITHOUT one. (A naive max(0, header extent − top) without the negative-top
    // guard would wrongly reserve the whole header here, pushing the body down.)
    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const negTall = await renderPage0(docWithHeader(body(), tallHeader, { marginTop: -10 }));
    const negNone = await renderPage0(docWithHeader(body(), null, { marginTop: -10 }));

    expect(negTall.filter((c) => c.text === 'HDR').length).toBeGreaterThan(0);
    expect(Math.min(...bodyY(negTall))).toBeCloseTo(Math.min(...bodyY(negNone)), 1);
  });
});
