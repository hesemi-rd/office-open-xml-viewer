import { describe, it, expect } from 'vitest';
import { computePages, paginateDocument } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, SectionProps, DocxDocumentModel,
  SectionGeom, PaginatedBodyElement,
} from './types';

// ECMA-376 §17.6.13 `<w:pgSz>` + §17.6.11 `<w:pgMar>` — page geometry is
// PER-SECTION. A mid-body SectionBreak carries its ending section's `geom`; the
// paginator stamps each element's `sectionGeom` (upcoming SectionBreak's geom, or
// the body-level section for the final section) so the renderer sizes each page
// from its own section. Single-section documents stamp the body-level geometry
// everywhere (byte-identical layout).

// Deterministic stub canvas: glyph advance = charCount × fontPx, font box =
// 0.8/0.2 em (a single line is exactly fontPx tall). Copied from pagination.test.ts.
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

type DocRun = DocParagraph['runs'][number];
function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function para(text: string, fontSize = 20): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [textRun(text, fontSize)],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

// A portrait first section (200×140) ended by a nextPage break, then a landscape
// body section (140×200) via doc.section. `geom` on the break carries the portrait
// section's page size; doc.section carries the landscape body size.
function mixedDoc(): DocxDocumentModel {
  const portrait: SectionGeom = {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0,
  };
  const bodySection: SectionProps = {
    pageWidth: 140, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  const body: BodyElement[] = [
    para('PORTRAIT_SECTION'),
    { type: 'sectionBreak', kind: 'nextPage', geom: portrait } as BodyElement,
    para('LANDSCAPE_SECTION'),
  ];
  return {
    section: bodySection, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('per-section page geometry (§17.6.13/§17.6.11) — paginator', () => {
  it('stamps each element with its section geometry', () => {
    const doc = mixedDoc();
    const pages = computePages(doc.body, doc.section, makeCtx());
    // Page 0 = portrait section (the first section, ended by the nextPage break).
    const p0 = pages[0].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(p0.sectionGeom?.pageWidth).toBe(200);
    expect(p0.sectionGeom?.pageHeight).toBe(140);
    // Page 1 = landscape body section (no following break ⇒ body-level geometry).
    const p1 = pages[1].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(p1.sectionGeom?.pageWidth).toBe(140);
    expect(p1.sectionGeom?.pageHeight).toBe(200);
  });

  it('single-section document stamps the body-level geometry on every element', () => {
    const section: SectionProps = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section, body: [para('A'), para('B')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = computePages(doc.body, doc.section, makeCtx());
    for (const page of pages) {
      for (const el of page) {
        expect((el as PaginatedBodyElement).sectionGeom?.pageWidth).toBe(200);
        expect((el as PaginatedBodyElement).sectionGeom?.pageHeight).toBe(140);
      }
    }
  });
});
