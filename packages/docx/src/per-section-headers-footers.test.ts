import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxDocumentModel, SectionProps, HeaderFooter, HeadersFooters,
} from './types';

// ECMA-376 §17.10.1 — per-section headers/footers + titlePage. A `SectionBreak`
// marker carries its ENDING section's resolved header/footer set + `<w:titlePg>`
// flag; the body-level section's set lives on `doc.headers`/`doc.footers`/
// `doc.section.titlePage`. The renderer must, per page, select the header/footer
// of the section ACTIVE at that page's top, applying the §17.10.1 precedence
// (first → even → default) with `first` keyed off the section's FIRST page (not
// document page 0). This guards the sample-13 bug: the title section's first-page
// footer (the DOI line) was dropped when a later section overwrote the global set.

// A minimal canvas that records every fillText call so a test can assert which
// header/footer text rendered on a given page. measureText returns a width
// proportional to length so layout produces non-degenerate lines.
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; texts: string[] } {
  let font = '10px serif';
  const texts: string[] = [];
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
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string) { texts.push(s); },
    strokeText(s: string) { texts.push(s); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, texts };
}

function para(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 11, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function footer(text: string): HeaderFooter {
  return { body: [para(text) as unknown as BodyElement] };
}

function hf(parts: Partial<HeadersFooters>): HeadersFooters {
  return { default: null, first: null, even: null, ...parts };
}

// Two sections, each a full page of its own. Section A ends at a nextPage section
// break carrying footerA (first + default) and titlePage=true. Section B is the
// final (body-level) section with footerB as its default. With pageHeight tuned so
// each paragraph fills exactly one page, page 0 belongs to section A, page 1 to B.
function twoSectionDoc(): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 120,
    marginTop: 10, marginRight: 10, marginBottom: 30, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  const body: BodyElement[] = [
    para('SECTION_A_BODY') as unknown as BodyElement,
    {
      type: 'sectionBreak', kind: 'nextPage',
      headers: hf({}),
      footers: hf({ first: footer('FOOTER_A_FIRST'), default: footer('FOOTER_A_DEFAULT') }),
      titlePage: true,
    } as BodyElement,
    para('SECTION_B_BODY') as unknown as BodyElement,
  ];
  return {
    section,
    body,
    headers: hf({}),
    // body-level (section B) default footer
    footers: hf({ default: footer('FOOTER_B_DEFAULT') }),
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderPageTexts(doc: DocxDocumentModel, pageIndex: number): Promise<string[]> {
  const { canvas, texts } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, pageIndex, { dpr: 1, width: 400 });
  return texts;
}

describe('per-section headers/footers (§17.10.1)', () => {
  it('page 0 = section A first page → renders section A first-page footer', async () => {
    const texts = await renderPageTexts(twoSectionDoc(), 0);
    const joined = texts.join('|');
    expect(joined).toContain('SECTION_A_BODY');
    // titlePage on section A + page 0 is its first page ⇒ `first` footer.
    expect(joined).toContain('FOOTER_A_FIRST');
    expect(joined).not.toContain('FOOTER_A_DEFAULT');
    // Section B's footer must NOT appear on page 0.
    expect(joined).not.toContain('FOOTER_B_DEFAULT');
  });

  it('page 1 = section B → renders section B default footer, not section A footer', async () => {
    const texts = await renderPageTexts(twoSectionDoc(), 1);
    const joined = texts.join('|');
    expect(joined).toContain('SECTION_B_BODY');
    expect(joined).toContain('FOOTER_B_DEFAULT');
    expect(joined).not.toContain('FOOTER_A_FIRST');
    expect(joined).not.toContain('FOOTER_A_DEFAULT');
  });

  it('falls back to body-level footer for a single-section document (unchanged path)', async () => {
    const section: SectionProps = {
      pageWidth: 400, pageHeight: 200,
      marginTop: 10, marginRight: 10, marginBottom: 30, marginLeft: 10,
      headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section, body: [para('ONLY_BODY') as unknown as BodyElement],
      headers: hf({}), footers: hf({ default: footer('GLOBAL_FOOTER') }),
      fontFamilyClasses: { 'Times New Roman': 'roman' },
    } as unknown as DocxDocumentModel;
    const texts = await renderPageTexts(doc, 0);
    const joined = texts.join('|');
    expect(joined).toContain('ONLY_BODY');
    expect(joined).toContain('GLOBAL_FOOTER');
  });
});
