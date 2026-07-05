import { describe, it, expect } from 'vitest';
import { computePages, renderDocumentToCanvas } from './renderer.js';
import { computePageNumbering } from './page-numbering.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, FieldRun, SectionProps, DocxDocumentModel,
  SectionGeom, PageNumType, PaginatedBodyElement, HeaderFooter,
} from './types';

// ECMA-376 §17.6.12 `<w:pgNumType>` — end-to-end coverage of per-section page-number
// RESTART (`w:start`) + FORMAT (`w:fmt`) and the §17.16.4.3.1 field `\*` switch,
// from the parsed model through pagination (which stamps `sectionPageNumType`) to
// the PAGE field text a footer paints. Mirrors per-section-page-geometry.test.ts's
// deterministic stub-canvas approach (glyph advance = charCount × fontPx, line
// height = fontPx) so pagination is exact and headless.

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

// OffscreenCanvas polyfill (paginateWithHeaderFooterReserve builds its measure ctx
// from `new OffscreenCanvas`, absent in node). Same deterministic stub.
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function pageFieldRun(instruction = 'PAGE', fontSize = 20): DocRun {
  const f: FieldRun = {
    fieldType: 'page', instruction, fallbackText: '?',
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', background: null, vertAlign: null,
  } as unknown as FieldRun;
  return { type: 'field', ...f } as DocRun;
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
function footerWithPageField(instruction = 'PAGE'): HeaderFooter {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [pageFieldRun(instruction, 20)],
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { body: [{ type: 'paragraph', ...p } as BodyElement] };
}

const GEOM = (): SectionGeom => ({
  pageWidth: 200, pageHeight: 140,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0,
});
// content height 100 ⇒ five 20pt lines per page.

/** A 2-section doc: front matter (fmt/start on the mid-body break) + body (final
 *  section, its pgNumType on doc.section). `bodyLines` per section controls how
 *  many physical pages each spans. */
function twoSectionDoc(
  frontPgNum: PageNumType | null,
  bodyPgNum: PageNumType | null,
  frontLines = 6,
  bodyLines = 6,
  footerInstr = 'PAGE',
): DocxDocumentModel {
  const front: BodyElement[] = [];
  for (let i = 0; i < frontLines; i++) front.push(para(`F${i}`));
  const body: BodyElement[] = [];
  for (let i = 0; i < bodyLines; i++) body.push(para(`B${i}`));
  const footer = footerWithPageField(footerInstr);
  const bodySection: SectionProps = {
    ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    pageNumType: bodyPgNum,
  } as unknown as SectionProps;
  return {
    section: bodySection,
    body: [
      ...front,
      {
        type: 'sectionBreak', kind: 'nextPage', geom: GEOM(),
        headers: { default: null, first: null, even: null },
        footers: { default: footer, first: null, even: null },
        titlePage: false,
        pageNumType: frontPgNum,
      } as BodyElement,
      ...body,
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: footer, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

/** A 2-section doc joined by a CONTINUOUS section break (§17.18.77): the second
 *  section begins MID-PAGE below the first (no page break) and its content SPILLS
 *  onto the following physical pages. The second section is the FINAL (body-level)
 *  section, so its `w:type="continuous"` lives on `section.sectionStart` and its
 *  `<w:pgNumType>` on `section.pageNumType` — exactly how the parser models
 *  sample-27 (§17.6.12). `firstLines` keeps section 1 short so section 2 shares its
 *  page; `secondLines` controls how far section 2 spills. */
function continuousSpilloverDoc(
  secondPgNum: PageNumType | null,
  firstLines: number,
  secondLines: number,
): DocxDocumentModel {
  const front: BodyElement[] = [];
  for (let i = 0; i < firstLines; i++) front.push(para(`S1-${i}`));
  const body: BodyElement[] = [];
  for (let i = 0; i < secondLines; i++) body.push(para(`S2-${i}`));
  const footer = footerWithPageField('PAGE');
  const bodySection: SectionProps = {
    ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    // §17.18.77 — the body-level (final) section starts CONTINUOUS, so it shares the
    // page where section 1 ends. §17.6.12 — and it carries the restart.
    sectionStart: 'continuous',
    pageNumType: secondPgNum,
  } as unknown as SectionProps;
  return {
    section: bodySection,
    body: [
      ...front,
      {
        // The mid-body marker ENDS section 1 (which carries no restart). The break's
        // effective kind is read from the UPCOMING section (the body section's
        // `sectionStart: 'continuous'`), so this marker's own `kind` is irrelevant.
        type: 'sectionBreak', kind: 'nextPage', geom: GEOM(),
        headers: { default: null, first: null, even: null },
        footers: { default: footer, first: null, even: null },
        titlePage: false,
        pageNumType: null,
      } as BodyElement,
      ...body,
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: footer, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('page-number restart + format — pagination stamp → computePageNumbering', () => {
  it('front matter lowerRoman(start=1) + body decimal(restart start=1)', () => {
    // 6 front lines / 6 body lines, 5 lines per page ⇒ 2 pages each ⇒ 4 physical.
    const doc = twoSectionDoc(
      { fmt: 'lowerRoman', start: 1 },
      { fmt: 'decimal', start: 1 },
    );
    const pages = computePages(doc.body, doc.section, makeCtx());
    expect(pages.length).toBe(4);
    const nums = computePageNumbering(pages);
    // §17.6.12: front matter i, ii; body RESTARTS to 1, 2.
    expect(nums).toEqual([
      { displayNumber: 1, format: 'lowerRoman' },
      { displayNumber: 2, format: 'lowerRoman' },
      { displayNumber: 1, format: 'decimal' },
      { displayNumber: 2, format: 'decimal' },
    ]);
  });

  it('front matter with no pgNumType numbers 1..N in decimal (byte-identical baseline)', () => {
    const doc = twoSectionDoc(null, null);
    const pages = computePages(doc.body, doc.section, makeCtx());
    const nums = computePageNumbering(pages).map((n) => `${n.displayNumber}/${n.format}`);
    expect(nums).toEqual(['1/decimal', '2/decimal', '3/decimal', '4/decimal']);
  });

  // ECMA-376 §17.6.12 + §17.18.77 — CONTINUOUS restart semantics (issue #804).
  // GROUND TRUTH: sample-27 (section 1 → continuous break + `w:pgNumType w:start="50"`
  // → section 2 spilling to later pages, footer PAGE) rendered in real Word prints
  // footers [Page 1, Page 51, Page 52]. The restart's series counts from the FIRST
  // physical page the section's content appears on — the SHARED page (page 0), which
  // it does not OWN (page 0's top displays section 1's number 1). So section 2's
  // series is 50 on the shared page 0, 51 on page 1, 52 on page 2. The pre-#804
  // implementation restarted at the SPILLOVER page (where section 2 first owns a top)
  // and produced [1, 50, 51] — one short, because it did not count the shared page as
  // the section's first page. This deterministic doc mirrors sample-27's shape:
  //   5 lines/page; section 1 = 3 lines ⇒ 2 lines of section 2 share page 0;
  //   section 2 = 9 lines ⇒ 2 on page 0, 5 on page 1, 2 on page 2 ⇒ 3 pages.
  it('continuous restart counts the shared page as the section first page (§17.6.12, #804)', () => {
    const doc = continuousSpilloverDoc({ start: 50 }, 3, 9);
    const pages = computePages(doc.body, doc.section, makeCtx());
    expect(pages.length).toBe(3);
    // Section 2 shares page 0 (does not own its top) and owns pages 1–2.
    const nums = computePageNumbering(pages).map((n) => n.displayNumber);
    expect(nums).toEqual([1, 51, 52]);
  });

  // A continuous restart section that is a MID-PAGE ISLAND — its content fits within
  // the page it starts on and NEVER owns a page top (the next section takes over the
  // following page's top). Its `w:start` therefore never surfaces as a displayed
  // number; numbering stays sequential. No real sample has this shape (probed in the
  // browser, sample-13's `start=2` continuous section begins exactly at a page
  // boundary and OWNS that page's top with anchor offset 0 — see the module header
  // of page-numbering.ts), so this arm is pinned here deterministically.
  it('a continuous restart island that owns no page top does not surface its start', () => {
    // section 1 = 2 lines; section 2 (start=99) = 2 lines ⇒ both share page 0 (4 of 5
    // lines). Section 2 owns no page top (it is the final section here, but its whole
    // content sits on page 0). Single page ⇒ display 1, no restart visible.
    const doc = continuousSpilloverDoc({ start: 99 }, 2, 2);
    const pages = computePages(doc.body, doc.section, makeCtx());
    expect(pages.length).toBe(1);
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([1]);
  });
});

// Recording canvas — captures fillText so we can read the exact PAGE-field string
// painted in each page's footer.
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; texts: string[] } {
  let font = '10px serif';
  const texts: string[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p,
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
    fillText(s: string) { texts.push(s); }, strokeText(s: string) { texts.push(s); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, texts };
}

describe('PAGE field renders the per-section displayed number (footer)', () => {
  async function footerTexts(doc: DocxDocumentModel, pageIndex: number): Promise<string[]> {
    const { canvas, texts } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, pageIndex, { dpr: 1 });
    return texts;
  }

  it('paints i, ii, then restarts to 1, 2 across the two sections', async () => {
    const doc = twoSectionDoc(
      { fmt: 'lowerRoman', start: 1 },
      { fmt: 'decimal', start: 1 },
    );
    // The footer's PAGE field is the only glyph on each page besides body text; the
    // roman/decimal string is unique enough to assert via `.includes`.
    expect(await footerTexts(doc, 0)).toContain('i');
    expect(await footerTexts(doc, 1)).toContain('ii');
    expect(await footerTexts(doc, 2)).toContain('1');
    expect(await footerTexts(doc, 3)).toContain('2');
  });

  it('start=0 offsets the first page number to 0 (Word writes start="0")', async () => {
    const doc = twoSectionDoc({ start: 0 }, null, 6, 1);
    // page 0 shows "0", page 1 shows "1" (decimal — no fmt on the front section).
    expect(await footerTexts(doc, 0)).toContain('0');
    expect(await footerTexts(doc, 1)).toContain('1');
  });

  it('field \\* switch overrides the section fmt (PAGE \\* Roman on a decimal section)', async () => {
    // Body section decimal; the footer field carries `\* Roman` ⇒ uppercase roman.
    const doc = twoSectionDoc(null, { fmt: 'decimal', start: 3 }, 1, 1, 'PAGE \\* Roman');
    // front page 0 = 1 (decimal, no fmt) but its footer also carries \* Roman ⇒ "I".
    expect(await footerTexts(doc, 0)).toContain('I');
    // body page 1 restarts to 3; \* Roman ⇒ "III".
    expect(await footerTexts(doc, 1)).toContain('III');
  });

  it('single-section document without pgNumType is unchanged (decimal 1..N)', async () => {
    const footer = footerWithPageField('PAGE');
    const section: SectionProps = {
      ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    } as unknown as SectionProps;
    const doc = {
      section,
      body: [para('A0'), para('A1'), para('A2'), para('A3'), para('A4'), para('A5')],
      headers: { default: null, first: null, even: null },
      footers: { default: footer, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    // 6 lines, 5 per page ⇒ 2 pages: footers "1" and "2".
    expect(await footerTexts(doc, 0)).toContain('1');
    expect(await footerTexts(doc, 1)).toContain('2');
  });
});
