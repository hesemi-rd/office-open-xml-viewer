import { describe, it, expect } from 'vitest';
import {
  computePages,
  paginateDocument,
  renderDocumentToCanvas,
  physicalPageSizeForPage,
  type DocxTextRunInfo,
} from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, SectionProps, DocxDocumentModel,
  SectionGeom, PaginatedBodyElement,
} from './types';

// ECMA-376 §17.6.20 — text direction is PER-SECTION (issue #1000, #988 batch-3
// adjudication ①): a vertical (tbRl ≡ btLr per Word GT) NON-FINAL section can
// coexist with a horizontal FINAL section in one document. The parser carries
// the ending section's `<w:textDirection>` on its SectionBreak marker; the
// paginator lays a vertical section out in its SWAPPED LOGICAL geometry
// (`verticalLayoutSection`'s quarter-turn: logical width = physical height,
// margins rotated) and stamps `sectionGeom` (logical) + `sectionTextDirection`
// in lockstep; the renderer rotates each page +90° per ITS OWN direction.

// Deterministic stub canvas: glyph advance = charCount × fontPx, font box =
// 0.8/0.2 em (a single line is exactly fontPx tall). Copied from
// per-section-page-geometry.test.ts.
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

// `paginateDocument` builds its measure ctx from `new OffscreenCanvas(...)`,
// which the node test env lacks — polyfill with the same deterministic stub.
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

const EMPTY_HF = { default: null, first: null, even: null };

// Horizontal FINAL (body) section with fully asymmetric margins so any margin
// rotation slip is caught.
function horizontalBodySection(): SectionProps {
  return {
    pageWidth: 400, pageHeight: 500,
    marginTop: 12, marginRight: 24, marginBottom: 36, marginLeft: 48,
    headerDistance: 5, footerDistance: 6, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

// A vertical section's PHYSICAL page geometry (as the parser emits it — the
// sectPr's verbatim pgSz/pgMar), asymmetric margins.
const VERT_PHYS: SectionGeom = {
  pageWidth: 612, pageHeight: 792,
  marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
  headerDistance: 7, footerDistance: 8,
};

// body: [V1] <break: vertical btLr, explicit geom> [H1] (horizontal body section).
function mixedDoc(): DocxDocumentModel {
  const body: BodyElement[] = [
    para('V1'),
    { type: 'sectionBreak', kind: 'nextPage', geom: VERT_PHYS, textDirection: 'btLr' } as BodyElement,
    para('H1'),
  ];
  return {
    section: horizontalBodySection(), body,
    headers: EMPTY_HF, footers: EMPTY_HF,
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('per-section text direction (§17.6.20, issue #1000) — paginator stamps', () => {
  it('stamps a vertical mid-section with its SWAPPED logical geometry + direction', () => {
    const doc = mixedDoc();
    const pages = computePages(doc.body, doc.section, makeCtx());
    // Page 0 = the btLr section. Its stamped geometry is the LOGICAL quarter-turn
    // of the marker's physical geom: logical width = physical height, and
    // logical margin{Left,Top,Right,Bottom} = physical margin{Top,Right,Bottom,Left}
    // (verticalLayoutSection). header/footer distances are preserved.
    const v1 = pages[0].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(v1.sectionTextDirection).toBe('btLr');
    expect(v1.sectionGeom).toEqual({
      pageWidth: 792, pageHeight: 612,
      marginLeft: 10, marginTop: 20, marginRight: 30, marginBottom: 40,
      headerDistance: 7, footerDistance: 8,
    });
    // Page 1 = horizontal final section: body-level geometry, direction null.
    const h1 = pages[1].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(h1.sectionTextDirection).toBeNull();
    expect(h1.sectionGeom?.pageWidth).toBe(400);
    expect(h1.sectionGeom?.pageHeight).toBe(500);
    expect(h1.sectionGeom?.marginLeft).toBe(48);
  });

  it('a geom-less vertical break swaps the BODY geometry (§17.18.77 inheritance)', () => {
    const doc = mixedDoc();
    doc.body[1] = { type: 'sectionBreak', kind: 'nextPage', textDirection: 'tbRl' } as BodyElement;
    const pages = computePages(doc.body, doc.section, makeCtx());
    const v1 = pages[0].find((e) => e.type === 'paragraph') as PaginatedBodyElement;
    expect(v1.sectionTextDirection).toBe('tbRl');
    // swap of body 400×500 / margins T12 R24 B36 L48 → logical
    // 500×400 / L12 T24 R36 B48.
    expect(v1.sectionGeom).toEqual({
      pageWidth: 500, pageHeight: 400,
      marginLeft: 12, marginTop: 24, marginRight: 36, marginBottom: 48,
      headerDistance: 5, footerDistance: 6,
    });
  });

  it('all-vertical single-section document keeps the legacy swapped stamps', () => {
    const section: SectionProps = {
      ...horizontalBodySection(),
      pageWidth: 612, pageHeight: 792,
      marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
      headerDistance: 7, footerDistance: 8,
      textDirection: 'tbRl',
    } as SectionProps;
    const doc = {
      section, body: [para('V1'), para('V2')],
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    // paginateDocument applies verticalLayoutDoc (the body-level swap) itself.
    const pages = paginateDocument(doc);
    for (const page of pages) {
      for (const el of page) {
        const p = el as PaginatedBodyElement;
        expect(p.sectionTextDirection).toBe('tbRl');
        expect(p.sectionGeom?.pageWidth).toBe(792);
        expect(p.sectionGeom?.pageHeight).toBe(612);
        expect(p.sectionGeom?.marginLeft).toBe(10);
      }
    }
  });

  // Height- AND width-sensitive: the vertical section must paginate against its
  // SWAPPED frame. Physical 200(w)×300(h), margins all 20 ⇒ logical page 300×200,
  // content height 160 (8 × 20pt one-line paras), wrap width 260 (a 10-char 20pt
  // para = 200pt fits on ONE line). If the frame were left physical: content
  // height 260 and wrap width 160 (two lines per para) — a totally different
  // distribution — so this fails the moment the swap is dropped.
  it('paginates a vertical mid-section against its SWAPPED logical frame', () => {
    const vertPhys: SectionGeom = {
      pageWidth: 200, pageHeight: 300,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0,
    };
    const body: BodyElement[] = [
      ...Array.from({ length: 10 }, (_v, i) => para(`A${i + 1}______A${i + 1}`.slice(0, 10))),
      { type: 'sectionBreak', kind: 'nextPage', geom: vertPhys, textDirection: 'tbRl' } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: horizontalBodySection(), body,
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = computePages(doc.body, doc.section, makeCtx());
    const counts = pages.map((page) => page.filter((e) => e.type === 'paragraph').length);
    // 8 one-line paras on page 0, the remaining 2 on page 1, B1 on page 2.
    expect(counts).toEqual([8, 2, 1]);
  });

  it('promotes a CONTINUOUS boundary to a page break when the direction changes', () => {
    // The paint model rotates a whole page: one direction per physical page is a
    // documented renderer constraint, so a continuous break whose flow direction
    // changes must open a fresh page. (Word GT for this boundary is unverified —
    // this pins the model constraint, not a Word behavior claim.)
    const section: SectionProps = { ...horizontalBodySection(), sectionStart: 'continuous' } as SectionProps;
    const body: BodyElement[] = [
      para('V1'),
      { type: 'sectionBreak', kind: 'nextPage', geom: VERT_PHYS, textDirection: 'btLr' } as BodyElement,
      para('H1'),
    ];
    const doc = {
      section, body, headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = computePages(doc.body, doc.section, makeCtx());
    expect(pages.length).toBe(2);

    // Control: the same continuous boundary WITHOUT a direction change stays on
    // one page (the existing §17.18.77 continuous behavior is untouched).
    const flatBody: BodyElement[] = [
      para('S1'),
      { type: 'sectionBreak', kind: 'nextPage', geom: VERT_PHYS } as BodyElement,
      para('S2'),
    ];
    const flatDoc = {
      section, body: flatBody, headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const flatPages = computePages(flatDoc.body, flatDoc.section, makeCtx());
    expect(flatPages.length).toBe(1);
  });
});

// A recording canvas that captures ctx.rotate + canvas dims + text-run overlay
// payloads, so a test can assert the per-page +90° rotation.
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; rotations: number[] } {
  let font = '10px serif';
  const rotations: number[] = [];
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
    rotate(r: number) { rotations.push(r); },
    setTransform() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, rotations };
}

describe('per-section text direction (§17.6.20, issue #1000) — renderer', () => {
  it('rotates ONLY the vertical page +90°; the horizontal page stays unrotated', async () => {
    const doc = mixedDoc();
    const runs0: DocxTextRunInfo[] = [];
    const { canvas: c0, rotations: r0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { dpr: 1, onTextRun: (r) => runs0.push(r) });
    const runs1: DocxTextRunInfo[] = [];
    const { canvas: c1, rotations: r1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { dpr: 1, onTextRun: (r) => runs1.push(r) });

    // Page 0 (btLr): the whole-page +90° paint rotation fires (one +π/2 among the
    // rotations; per-glyph counter-rotations may add more entries).
    expect(r0.some((r) => Math.abs(r - Math.PI / 2) < 1e-9)).toBe(true);
    // Its canvas is the PHYSICAL page box (612×792 portrait).
    expect(c0.width / c0.height).toBeCloseTo(612 / 792, 2);
    // The overlay run is handed the vertical placement transform.
    expect(runs0.length).toBeGreaterThan(0);
    expect(runs0[0].transform).toBeDefined();

    // Page 1 (horizontal body): no page rotation, no overlay transform, its own
    // 400×500 box.
    expect(r1.some((r) => Math.abs(r - Math.PI / 2) < 1e-9)).toBe(false);
    expect(c1.width / c1.height).toBeCloseTo(400 / 500, 2);
    expect(runs1.length).toBeGreaterThan(0);
    expect(runs1[0].transform).toBeUndefined();
  });
});

describe('per-section text direction (§17.6.20, issue #1000) — physical page size', () => {
  it('un-swaps each page by ITS OWN stamped direction', () => {
    const doc = mixedDoc();
    const pages = paginateDocument(doc);
    // Page 0 (vertical): stamped LOGICAL 792×612 → physical 612×792.
    expect(physicalPageSizeForPage(pages, 0, doc.section)).toEqual({ widthPt: 612, heightPt: 792 });
    // Page 1 (horizontal): stamped physical box verbatim.
    expect(physicalPageSizeForPage(pages, 1, doc.section)).toEqual({ widthPt: 400, heightPt: 500 });
  });

  it('resolves an EMPTY parity page from its page metadata, not the body fallback', () => {
    // Mid VERTICAL section whose start type is oddPage: the parity padding leaves
    // a BLANK page owned by the incoming vertical section (pages: [H1][blank][V1]
    // [H2]). Element-stamp fallback would report the body-level 400×500 box for
    // the blank page; the paginator's page metadata reports the vertical
    // section's physical 612×792.
    const body: BodyElement[] = [
      para('H1'),
      // marker ending section 1 (horizontal, geom-less). The UPCOMING section's
      // start type is the NEXT marker's kind (§17.6.22): oddPage.
      { type: 'sectionBreak', kind: 'nextPage' } as BodyElement,
      para('V1'),
      { type: 'sectionBreak', kind: 'oddPage', geom: VERT_PHYS, textDirection: 'tbRl' } as BodyElement,
      para('H2'),
    ];
    const doc = {
      section: horizontalBodySection(), body,
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = paginateDocument(doc);
    // Page 0: H1. Page 1: parity blank (the oddPage section pads past the even
    // slot). Page 2: V1 (vertical). Page 3: H2.
    expect(pages.length).toBe(4);
    expect(pages[1]).toHaveLength(0);
    expect(physicalPageSizeForPage(pages, 2, doc.section)).toEqual({ widthPt: 612, heightPt: 792 });
    // The blank parity page belongs to the INCOMING vertical section (the
    // paginator's transition model): physical 612×792, not the body 400×500.
    expect(physicalPageSizeForPage(pages, 1, doc.section)).toEqual({ widthPt: 612, heightPt: 792 });
  });
});

describe('per-section text direction (§17.6.20, issue #1000) — HF reserve', () => {
  it('reserves NOTHING on a vertical page (paint draws its HF physically, reserve-free)', () => {
    // Vertical mid-section: physical 200×300, margins 20 ⇒ logical content height
    // 160 ⇒ exactly 8 one-line 20pt paras. A 40pt footer whose extent (10+40)
    // overflows the 20pt logical bottom margin would phantom-reserve 30pt and
    // push A7/A8 off the page — the paint pass never reserves on vertical pages
    // (issue #988: vertical HF draw horizontally at the physical margins), so
    // pagination must not either.
    const vertPhys: SectionGeom = {
      pageWidth: 200, pageHeight: 300,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 10,
    };
    const footer40 = { body: [para('F', 40)] };
    const body: BodyElement[] = [
      ...Array.from({ length: 8 }, (_v, i) => para(`A${i + 1}`)),
      {
        type: 'sectionBreak', kind: 'nextPage', geom: vertPhys, textDirection: 'tbRl',
        headers: EMPTY_HF, footers: { default: footer40, first: null, even: null },
        titlePage: false,
      } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: { ...horizontalBodySection(), footerDistance: 10 },
      body,
      headers: EMPTY_HF, footers: { default: footer40, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = paginateDocument(doc);
    const textsOn = (i: number) =>
      (pages[i] ?? []).filter((e) => e.type === 'paragraph')
        .map((e) => ((e as { runs?: { text?: string }[] }).runs ?? []).map((r) => r.text).join(''));
    expect(textsOn(0)).toEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8']);
    expect(textsOn(1)).toEqual(['B1']);
  });
});
