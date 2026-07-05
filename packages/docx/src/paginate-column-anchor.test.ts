import { describe, expect, it } from 'vitest';
import { computePages } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  ImageRun,
  SectionProps,
} from './types.js';

// Paginator-level pin for ECMA-376 §20.4.3.4 `relativeFrom="column"` anchors:
// the MEASURE pass (computePages → buildMeasureState) must resolve a body-level
// column anchor against the SAME text-column band as the paint pass.
//
// Regression this pins (PR #844 review F1): buildMeasureState seeded
// `contentX: 0` while the paint pass seeds `contentX = marginLeft × scale`.
// That seed was inert while xContainer never read contentX, but once `column`
// resolves against contentX/contentW, the measure pass placed body-level column
// anchors a full marginLeft LEFT of where the paint pass draws them — floats
// migrated into (or out of) the wrap band only during pagination, splitting
// paragraphs differently from the painted layout (sample-9 p3–10 drift).
//
// The MinState unit tests in anchor-column.test.ts cannot catch this: they
// hand-construct contentX and never exercise the paginator's own state seeding.
// This test goes through computePages so the real buildMeasureState seed is on
// the hook.
//
// Setup (mirrors page-anchor-prescan.test.ts geometry): 200×200 page, 20pt
// margins ⇒ text band x∈[20,180] (contentW=160), contentH=160. 20pt stub font ⇒
// 8 chars/line, 8 lines/page. Paragraph A carries 64 chars = exactly 8 full-width
// lines. Paragraph B carries a page-level square float: 80×60 at anchorXPt=160,
// anchorY margin+0 ⇒ y∈[20,80].
//
//   • CORRECT (§20.4.3.4, body level: column band == margin band, start=20):
//     float x∈[180,260] — fully OUTSIDE the text band ⇒ A keeps 8 chars/line,
//     8 lines fit page 1 exactly ⇒ NOT split (no lineSlice), identical to the
//     margin-relative control.
//   • BUGGY measure seed (contentX=0): float x∈[160,240] — overlaps the band's
//     right 20pt ⇒ the 3 lines in y∈[20,80] narrow to 140pt = 7 chars ⇒ 9 lines
//     total = 180pt > 160 ⇒ A splits onto page 2 (lineSlice appears).

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
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function section(): SectionProps {
  return {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  };
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(runs: DocRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** Page-level square wrap float, 80×60, X from `relativeFrom` + offset 160. */
function floatImage(anchorXRelativeFrom: string): DocRun {
  const img: ImageRun = {
    imagePath: 'word/media/test1.png',
    mimeType: 'image/png',
    widthPt: 80,
    heightPt: 60,
    anchor: true,
    anchorXPt: 160,
    anchorYPt: 0,
    anchorXFromMargin: true,
    anchorYFromPara: false,
    wrapMode: 'square',
    wrapSide: 'bothSides',
    anchorXRelativeFrom,
    anchorYRelativeFrom: 'margin',
  };
  return { type: 'image', ...img } as DocRun;
}

function pagesFor(anchorXRelativeFrom: string): ReturnType<typeof computePages> {
  const body: BodyElement[] = [
    para([textRun('あ'.repeat(64), 20)]),
    para([floatImage(anchorXRelativeFrom)]),
  ];
  return computePages(body, section(), makeCtx());
}

describe('computePages — body-level column anchor measures like the paint pass (§20.4.3.4)', () => {
  it('column-relative float paginates identically to the margin-relative control', () => {
    const marginPages = pagesFor('margin');
    const columnPages = pagesFor('column');

    type SlicedEl = { lineSlice?: { start: number; end: number } };
    const marginFirst = marginPages[0][0] as SlicedEl;
    const columnFirst = columnPages[0][0] as SlicedEl;

    // Control: margin-anchored float at x∈[180,260] is outside the text band ⇒
    // paragraph A's 8 lines fit page 1 whole — no split.
    expect(marginFirst.lineSlice).toBeUndefined();

    // Pin: at body level the column band == the margin band, so the column
    // variant must behave IDENTICALLY. Under the buggy contentX=0 measure seed
    // the float lands at x∈[160,240], invades the band, and A splits
    // (lineSlice appears) — a pagination divergence from the paint pass.
    expect(columnFirst.lineSlice).toBeUndefined();
    expect(columnPages.length).toBe(marginPages.length);
  });
});
