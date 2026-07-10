import { describe, it, expect, beforeAll } from 'vitest';
import { layoutDocument } from './document-layout.js';
import {
  fragmentLineAdvancesPt,
  paragraphFragmentAdvancePt,
  type ParagraphFragment,
} from './layout-fragments.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// PR 5 Task 12 — body layout fragments.
//
// `layoutDocument(doc)` produces an immutable `DocumentLayout`: pages of
// `PlacedFragment`s wrapping `ParagraphFragment`s. Each fragment references its
// SOURCE paragraph (never a mutated copy), a placement-aware `MeasuredParagraph`,
// and a `[lineStart, lineEnd)` line range. This suite pins the fragment model
// contract (design doc §"Measured Fragment Model" / §"Pagination and paint
// invariants"): source identity, immutable line ranges, placement coordinates,
// page geometry, section context, paragraph continuation across pages, and the
// spacing-ownership invariant
//   cursor advancement == leadingSpacePt + measured line advances + trailingSpacePt.
// ─────────────────────────────────────────────────────────────────────────────

/** OffscreenCanvas polyfill with a linear glyph metric (width = fontPx * 0.5),
 *  matching the other renderer suites so `layoutDocument`'s scale-1 measurement is
 *  deterministic in node. */
function makeStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeStubCtx(); }
  };
});

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

/** Every paragraph fragment on every page, in document order. */
function allFragments(model: DocxDocumentModel) {
  const layout = layoutDocument(model);
  return layout.pages.flatMap((page) =>
    page.fragments.map((placed) => ({ page, placed, fragment: placed.fragment as ParagraphFragment })),
  );
}

describe('layoutDocument — body paragraph fragments (PR 5 Task 12)', () => {
  it('emits one placed paragraph fragment per body paragraph, referencing the SOURCE paragraph', () => {
    const p1 = para('alpha');
    const p2 = para('beta');
    const model = doc([p1 as unknown as BodyElement, p2 as unknown as BodyElement]);
    const layout = layoutDocument(model);

    expect(layout.pages.length).toBe(1);
    const frags = layout.pages[0].fragments;
    expect(frags.length).toBe(2);
    // Source identity: the fragment points at the parsed paragraph, not a clone.
    expect(frags[0].fragment.source).toBe(p1);
    expect(frags[1].fragment.source).toBe(p2);
    expect(frags[0].fragment.kind).toBe('paragraph');
  });

  it('keeps fragment state OFF the source paragraph (no fragment fields added to DocParagraph)', () => {
    // The paginator's established PaginatedBodyElement runtime stamps (colIndex,
    // colGeom, sectionGeom, ...) are pre-existing and out of scope; this pins that
    // the NEW fragment model never writes its fields onto the parsed paragraph — the
    // measurement/line-range/spacing live in the layout result, keyed off-object.
    const p1 = para('gamma');
    layoutDocument(doc([p1 as unknown as BodyElement]));
    const record = p1 as unknown as Record<string, unknown>;
    for (const field of ['measured', 'lineStart', 'lineEnd', 'leadingSpacePt', 'trailingSpacePt', 'fragment', 'placedFragment']) {
      expect(record[field]).toBeUndefined();
    }
  });

  it('records immutable line ranges covering the whole paragraph', () => {
    const p = para(Array.from({ length: 40 }, () => 'w').join(' '));
    const layout = layoutDocument(doc([p as unknown as BodyElement]));
    const frag = layout.pages[0].fragments[0].fragment as ParagraphFragment;
    expect(frag.lineStart).toBe(0);
    expect(frag.lineEnd).toBe(frag.measured.lines.length);
    expect(frag.lineEnd).toBeGreaterThan(1); // actually wrapped
  });

  it('places fragments with page-absolute coordinates and the content-band width', () => {
    const p1 = para('one');
    const p2 = para('two');
    const layout = layoutDocument(doc([p1 as unknown as BodyElement, p2 as unknown as BodyElement]));
    const [f1, f2] = layout.pages[0].fragments;
    // contentX = marginLeft, width = pageWidth - marginLeft - marginRight.
    expect(f1.xPt).toBeCloseTo(10, 6);
    expect(f1.widthPt).toBeCloseTo(180, 6);
    expect(f1.columnIndex).toBe(0);
    // The first fragment starts at the top content inset; the second stacks below it.
    expect(f1.yPt).toBeCloseTo(10, 6);
    expect(f2.yPt).toBeGreaterThan(f1.yPt);
    expect(f2.yPt).toBeCloseTo(f1.yPt + f1.heightPt, 6);
  });

  it('exposes page geometry and the resolved section context', () => {
    const layout = layoutDocument(doc([para('x') as unknown as BodyElement]));
    const page = layout.pages[0];
    expect(page.pageIndex).toBe(0);
    expect(page.geometry.pageWidth).toBe(200);
    expect(page.geometry.pageHeight).toBe(400);
    expect(page.geometry.marginLeft).toBe(10);
    // SectionLayoutContext carries the resolved grid policy (docGrid absent => none).
    expect(page.section.grid.kind).toBe('none');
  });

  it('splits a long paragraph into continuation fragments over one source', () => {
    // Force several pages: short page height, a long wrapping paragraph.
    const p = para(Array.from({ length: 300 }, () => 'w').join(' '));
    const layout = layoutDocument(doc([p as unknown as BodyElement], 60));
    const frags = allFragments(doc([p as unknown as BodyElement], 60));

    expect(layout.pages.length).toBeGreaterThan(1);
    // All continuation fragments share ONE source paragraph.
    for (const f of frags) expect(f.fragment.source).toBe((frags[0].fragment).source);

    // Line ranges are disjoint and contiguous, covering every measured line.
    const totalLines = frags[0].fragment.measured.lines.length;
    let cursor = 0;
    for (const f of frags) {
      expect(f.fragment.lineStart).toBe(cursor);
      expect(f.fragment.lineEnd).toBeGreaterThan(f.fragment.lineStart);
      cursor = f.fragment.lineEnd;
    }
    expect(cursor).toBe(totalLines);

    // Leading spacing only on the first fragment; trailing only on the last.
    expect(frags[frags.length - 1].fragment.trailingSpacePt).toBeGreaterThanOrEqual(0);
    for (let i = 1; i < frags.length; i++) {
      expect(frags[i].fragment.leadingSpacePt).toBe(0);
    }
    for (let i = 0; i < frags.length - 1; i++) {
      expect(frags[i].fragment.trailingSpacePt).toBe(0);
    }
  });

  it('remeasures at a changed wrap width: a fragment measurement matches its own placement', () => {
    const p = para(Array.from({ length: 40 }, () => 'w').join(' '));
    const wide = layoutDocument(doc([p as unknown as BodyElement])); // width 180
    const narrowModel = doc([p as unknown as BodyElement]);
    (narrowModel.section as SectionProps).pageWidth = 120; // width 100
    const narrow = layoutDocument(narrowModel);

    const wf = wide.pages[0].fragments[0].fragment as ParagraphFragment;
    const nf = narrow.pages[0].fragments[0].fragment as ParagraphFragment;
    // Each fragment's measurement reflects its own available width (no stale reuse).
    expect(wf.measured.placement.availableWidthPt).toBeCloseTo(180, 6);
    expect(nf.measured.placement.availableWidthPt).toBeCloseTo(100, 6);
    // A narrower band wraps to more lines.
    expect(nf.measured.lines.length).toBeGreaterThan(wf.measured.lines.length);
  });

  it('INVARIANT: cursor advancement == leadingSpacePt + line advances + trailingSpacePt', () => {
    const p1 = para(Array.from({ length: 30 }, () => 'w').join(' '), { spaceBefore: 6, spaceAfter: 8 });
    const p2 = para('tail', { spaceBefore: 4, spaceAfter: 3 });
    const frags = allFragments(doc([p1 as unknown as BodyElement, p2 as unknown as BodyElement]));
    for (const { placed, fragment } of frags) {
      // Height the paginator charged equals the spacing-owned decomposition.
      expect(placed.heightPt).toBeCloseTo(paragraphFragmentAdvancePt(fragment), 6);
      // Line advances sum to the measured content span (no double count) for a full range.
      if (fragment.lineStart === 0 && fragment.lineEnd === fragment.measured.lines.length && !fragment.measured.markOnly) {
        expect(fragmentLineAdvancesPt(fragment)).toBeCloseTo(
          fragment.measured.contentEndYPt - fragment.measured.contentStartYPt, 6,
        );
        // Leading spacing equals the gap the measurement placed above the first line.
        expect(fragment.leadingSpacePt).toBeCloseTo(
          fragment.measured.contentStartYPt - fragment.measured.placement.startYPt, 6,
        );
      }
    }
  });

  it('freezes the layout result and its fragment arrays', () => {
    const layout = layoutDocument(doc([para('x') as unknown as BodyElement]));
    expect(Object.isFrozen(layout)).toBe(true);
    expect(Object.isFrozen(layout.pages)).toBe(true);
    expect(Object.isFrozen(layout.pages[0])).toBe(true);
    expect(Object.isFrozen(layout.pages[0].fragments)).toBe(true);
    expect(Object.isFrozen(layout.pages[0].fragments[0])).toBe(true);
    // The inner ParagraphFragment is frozen too — its measured lines, line range and
    // spacing are immutable, so paint can never mutate the layout result (design
    // §"Pagination and paint invariants" 4).
    expect(Object.isFrozen(layout.pages[0].fragments[0].fragment)).toBe(true);
  });
});
