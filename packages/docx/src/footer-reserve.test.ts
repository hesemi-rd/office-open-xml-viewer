import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, HeaderFooter, SectionProps } from './types';

// ECMA-376 §17.10.1 — a footer taller than the bottom-margin allowance
// (marginBottom − footerDistance) rises ABOVE the content area and would overlap
// body text. Word never lays body text over a footer: it reserves the overflow so
// content breaks to the next page instead. The paginator measures each page's
// footer and re-paginates with that reservation (paginateWithFooterReserve). These
// tests pin that invariant with a synthetic doc whose first-page footer is far
// taller than its bottom margin. Reconstructed from sample-13's masthead footer
// behaviour (a DOI / corresponding-author block ~53pt tall over a ~49pt margin).

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

// pageHeight 600, margins 10, footerDistance 4 → bottom-margin allowance for the
// footer is marginBottom − footerDistance = 6pt. A footer taller than 6pt overflows.
function docWithFooter(body: BodyElement[], footer: HeaderFooter | null): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 600,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage',
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: footer, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderPage0(doc: DocxDocumentModel): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('footer reserve — body content never overlaps a tall footer (ECMA-376 §17.10.1)', () => {
  // Enough single-line body paragraphs to overflow page 0 (content area is 580pt;
  // ~50 lines is well over a page) so, absent any reservation, body text packs all
  // the way down to the content bottom (600 − marginBottom = 590pt).
  const body = (): BodyElement[] =>
    Array.from({ length: 50 }, () => para('BODY') as unknown as BodyElement);
  // A footer far taller than the 6pt allowance (6 lines ≈ ~70pt), so its top edge
  // rises well above the content bottom and into the body's region.
  const tallFooter: HeaderFooter = {
    body: Array.from({ length: 6 }, () => para('FTR') as unknown as BodyElement),
  };

  it('breaks body to the next page so no line is painted into the tall footer band', async () => {
    const withFooter = await renderPage0(docWithFooter(body(), tallFooter));
    const noFooter = await renderPage0(docWithFooter(body(), null));

    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const footerY = (calls: Call[]) => calls.filter((c) => c.text === 'FTR').map((c) => c.y);

    const maxBodyTall = Math.max(...bodyY(withFooter));
    const minFooter = Math.min(...footerY(withFooter));
    const maxBodyNone = Math.max(...bodyY(noFooter));

    // Sanity: the tall footer is actually painted on page 0.
    expect(footerY(withFooter).length).toBeGreaterThan(0);

    // INVARIANT: with the tall footer reserved, the lowest body line on page 0 sits
    // ABOVE the topmost footer line — body never overlaps the footer.
    expect(maxBodyTall).toBeLessThan(minFooter);

    // NON-TRIVIALITY: the SAME body, with no footer to reserve against, packs down
    // PAST where the tall footer's top sits — proving the invariant above is not
    // satisfied trivially (the body genuinely reaches into that band) and that the
    // reservation, not a short body, is what clears the footer.
    expect(maxBodyNone).toBeGreaterThan(minFooter);
  });
});
