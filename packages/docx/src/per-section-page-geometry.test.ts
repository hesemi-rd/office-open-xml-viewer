import { describe, it, expect } from 'vitest';
import { computePages, paginateDocument, renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, SectionProps, DocxDocumentModel,
  SectionGeom, PaginatedBodyElement, HeaderFooter,
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

// `paginateDocument` builds its measure ctx from `new OffscreenCanvas(...)`, which
// the node test env lacks. Polyfill it with the SAME deterministic stub (line
// height = fontPx) so the HF-reserve paginator runs headless and its page-fit
// arithmetic is exact. (The paginator-stamp/renderer tests above inject their ctx
// directly via computePages/renderDocumentToCanvas and don't need this.)
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

  // Height-sensitive spill — proves the const→arrow conversion of the page frame.
  // A first section of 200×140 with margins 20 has a content height of 100pt ⇒ five
  // 20pt paragraphs per page, so a SIXTH paragraph spills to a second page WITHIN the
  // first section. If the frame were still read from the body-level section (140×200,
  // content 160 ⇒ eight 20pt lines fit), all six would stay on page 0 and this test
  // would fail — i.e. it regresses the moment the per-section frame reverts to a const.
  it('paginates each section against ITS OWN page height (const→arrow)', () => {
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
    // Six 20pt paragraphs in the portrait first section, then a nextPage break into
    // the landscape body section carrying one paragraph.
    const body: BodyElement[] = [
      para('A1'), para('A2'), para('A3'), para('A4'), para('A5'), para('A6'),
      { type: 'sectionBreak', kind: 'nextPage', geom: portrait } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: bodySection, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = computePages(doc.body, doc.section, makeCtx());
    const texts = pages.map((page) =>
      page
        .filter((e) => e.type === 'paragraph')
        .map((e) => (e as unknown as { runs: { text: string }[] }).runs.map((r) => r.text).join('')),
    );
    // Portrait content height 100 ⇒ A1..A5 on page 0, A6 spills to page 1 (still the
    // portrait section), then the break opens page 2 for the landscape section.
    expect(texts).toEqual([['A1', 'A2', 'A3', 'A4', 'A5'], ['A6'], ['B1']]);
    // The spilled A6 still carries the portrait section geometry (it precedes the
    // break); B1 carries the body-level landscape geometry.
    const a6 = pages[1].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(a6.sectionGeom?.pageWidth).toBe(200);
    expect(a6.sectionGeom?.pageHeight).toBe(140);
    const b1 = pages[2].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(b1.sectionGeom?.pageWidth).toBe(140);
    expect(b1.sectionGeom?.pageHeight).toBe(200);
  });

  // Geom-less middle break — exercises `e.geom ?? bodySectionGeom`. Three sections:
  // break1 carries geom, break2 does NOT. `sectionGeomFrom` walks FORWARD, so the
  // element BETWEEN the two breaks (S2) belongs to the section ENDING at break2,
  // which has no geom ⇒ it falls back to the body-level geometry. S1 (before break1)
  // gets break1's geom; S3 (after break2, the final section) gets the body geometry.
  it('falls back to body geometry for a section whose break carries no geom', () => {
    const geom1: SectionGeom = {
      pageWidth: 300, pageHeight: 400,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0,
    };
    const bodySection: SectionProps = {
      pageWidth: 140, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const body: BodyElement[] = [
      para('S1'),
      { type: 'sectionBreak', kind: 'nextPage', geom: geom1 } as BodyElement,
      para('S2'),
      // No `geom`: this section inherits pgSz/pgMar ⇒ bodySectionGeom fallback.
      { type: 'sectionBreak', kind: 'nextPage' } as BodyElement,
      para('S3'),
    ];
    const doc = {
      section: bodySection, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = computePages(doc.body, doc.section, makeCtx());
    const s1 = pages[0].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    const s2 = pages[1].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    const s3 = pages[2].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    // S1: section ending at break1 ⇒ break1's geom.
    expect(s1.sectionGeom?.pageWidth).toBe(300);
    expect(s1.sectionGeom?.pageHeight).toBe(400);
    // S2: section ending at break2, which has NO geom ⇒ body-level geometry.
    expect(s2.sectionGeom?.pageWidth).toBe(140);
    expect(s2.sectionGeom?.pageHeight).toBe(200);
    // S3: final section ⇒ body-level geometry.
    expect(s3.sectionGeom?.pageWidth).toBe(140);
    expect(s3.sectionGeom?.pageHeight).toBe(200);
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

// A single 40pt paragraph measures exactly 40pt tall under the mock ctx (line
// height = fontPx). Same footer shape as per-section-headers-footers.test.ts's
// `footer()` — { body: [para(...)] } — but sized to 40pt so the reserve arithmetic
// is round. Verified empirically: measureHeaderFooterHeight(FOOTER_40PT) = 40.
const FOOTER_40PT: HeaderFooter = { body: [para('F', 40)] };

describe('per-section HF reserves (§17.6.11) — paginator', () => {
  it('reserves footer overflow against the PAGE section margins, not the body section', () => {
    // Mid-body section sec1 (page 0): pageHeight 200, marginTop 20, marginBottom 60,
    // footerDistance 10 ⇒ footer extent 10+40=50 FITS the 60pt bottom margin ⇒ reserve
    // 0. Its frame = 200−20−60 = 120 ⇒ five 20pt lines (100pt) fit.
    // Body section (page 1): marginBottom 20 ⇒ footer extent 50 overflows ⇒ reserve 30
    // (its OWN page). B1 fits regardless.
    // WITHOUT the fix, computeFooterReserves reads the body-level marginBottom (20) for
    // EVERY page ⇒ page 0 gets a phantom reserve 30 ⇒ usable 90 ⇒ only 4 lines ⇒ A5
    // spills to page 1. WITH the fix, page 0 reads sec1's stamped marginBottom (60) ⇒
    // reserve 0 ⇒ all five A* stay on page 0.
    const sec1: SectionGeom = {
      pageWidth: 200, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 60, marginLeft: 20,
      headerDistance: 0, footerDistance: 10,
    };
    const bodySection: SectionProps = {
      pageWidth: 200, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 10, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section: bodySection,
      body: [
        para('A1'), para('A2'), para('A3'), para('A4'), para('A5'),
        {
          // The mid-body section must carry the footer itself — otherwise it resolves
          // no footer on its pages and the (phantom) reserve is never applied there,
          // so the bug can't surface. Mirror doc.footers on the break.
          type: 'sectionBreak', kind: 'nextPage', geom: sec1,
          headers: { default: null, first: null, even: null },
          footers: { default: FOOTER_40PT, first: null, even: null },
          titlePage: false,
        } as BodyElement,
        para('B1'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: FOOTER_40PT, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = paginateDocument(doc);
    const textsOn = (i: number) =>
      (pages[i] ?? []).filter((e) => e.type === 'paragraph')
        .map((e) => ((e as { runs?: { text?: string }[] }).runs ?? []).map((r) => r.text).join(''));
    // All five A-paragraphs fit page 0 (reserve 0 under ITS section's marginBottom 60).
    expect(textsOn(0)).toEqual(['A1', 'A2', 'A3', 'A4', 'A5']);
    expect(textsOn(1)).toEqual(['B1']);
  });
});

// A recording canvas that exposes width/height + records fillText, so a test can
// assert both the page pixel size and which paragraph text landed on a page.
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; texts: string[] } {
  let font = '10px serif';
  const texts: string[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
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
    fillText(s: string) { texts.push(s); }, strokeText(s: string) { texts.push(s); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, texts };
}

describe('per-section page geometry (§17.6.13/§17.6.11) — renderer', () => {
  it('sizes each page from its own section (portrait then landscape)', async () => {
    const doc = mixedDoc();
    // dpr:1 so canvas.width == cssWidth. No `width` override ⇒ each page sizes from
    // its OWN section's pageWidth × PT_TO_PX. Page 0 = portrait 200×140,
    // page 1 = landscape 140×200. The bug was every page sizing from doc.section
    // (140×200), so page 0 would come out 140-wide.
    const { canvas: c0, texts: t0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { dpr: 1 });
    const { canvas: c1, texts: t1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { dpr: 1 });

    // Join with '' — the deterministic mock ctx wraps "PORTRAIT_SECTION" into two
    // fillText calls at the content-width boundary, so the text arrives split.
    expect(t0.join('')).toContain('PORTRAIT_SECTION');
    expect(t1.join('')).toContain('LANDSCAPE_SECTION');
    // Portrait page is wider than tall; landscape page is taller than wide. This is
    // the load-bearing discriminator: before the fix page 0 sized from the body-level
    // (landscape 140×200) section, so c0 would be TALLER than wide (test Red).
    expect(c0.width).toBeGreaterThan(c0.height);
    expect(c1.height).toBeGreaterThan(c1.width);
    // Aspect ratios match the two sections (200/140 vs 140/200). Precision 2: canvas
    // dims are integer-pixel (Math.round(cssW/H*dpr)), so the ratio carries a ~7e-4
    // rounding residual — precision 2 still cleanly separates 1.43 from the buggy 0.70.
    expect(c0.width / c0.height).toBeCloseTo(200 / 140, 2);
    expect(c1.width / c1.height).toBeCloseTo(140 / 200, 2);
  });

  it('honours opts.width per CALL: same pixel width, per-page height from aspect', async () => {
    // The scroll viewer's contract: `width` is the canvas CSS width for THIS call.
    // Mixed-width pages rendered at a constant width fill the same pixels at
    // different px-per-pt scales; height still follows each page's OWN aspect.
    // Page 0 (portrait 200×140): height = 400·140/200 = 280.
    // Page 1 (landscape 140×200): height = 400·200/140 ≈ 571.
    const doc = mixedDoc();
    const { canvas: c0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { width: 400, dpr: 1 });
    const { canvas: c1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { width: 400, dpr: 1 });
    expect(c0.width).toBe(400);
    expect(c1.width).toBe(400);
    expect(c0.height).toBe(280);
    expect(c1.height).toBe(Math.round(400 * (200 / 140)));
  });

  it('single-section document sizes every page from doc.section (no regression)', async () => {
    const section: SectionProps = {
      pageWidth: 300, pageHeight: 400,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section, body: [para('ONLY_SECTION')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const { canvas } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1 });
    // Precision 2: integer-pixel rounding gives 400/533 = 0.7505, ~5e-4 off 0.75.
    expect(canvas.width / canvas.height).toBeCloseTo(300 / 400, 2);
  });
});

describe('page-size fact (§17.6.13/§17.6.11) — paginateDocument sectionGeom', () => {
  it('exposes per-page width/height via the first element sectionGeom', () => {
    // paginateDocument is the shared primitive DocxDocument.pageSize (main mode)
    // and render-worker's pageSizes[] both read. Each page's first element carries
    // its section geometry.
    const doc = mixedDoc();
    const pages = paginateDocument(doc);
    const sizeOf = (i: number) => {
      const g = (pages[i]?.[0] as PaginatedBodyElement | undefined)?.sectionGeom;
      return { widthPt: g?.pageWidth ?? doc.section.pageWidth, heightPt: g?.pageHeight ?? doc.section.pageHeight };
    };
    expect(sizeOf(0)).toEqual({ widthPt: 200, heightPt: 140 });
    expect(sizeOf(1)).toEqual({ widthPt: 140, heightPt: 200 });
  });
});
