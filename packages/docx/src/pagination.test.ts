import { describe, it, expect } from 'vitest';
import { computePages, computeColumns } from './renderer.js';
import type { BodyElement, DocParagraph, DocxTextRun, ShapeRun, SectionProps, PaginatedBodyElement } from './types';

// Unit tests for computePages pagination behaviour that the renderer-path VRT
// (local-only, private samples) cannot guard in CI. A deterministic stub canvas
// makes line wrapping and line heights predictable: glyph advance = charCount ×
// fontPx, and the font box = 0.8/0.2 em (so a single line is exactly fontPx tall
// with no spacing/grid). CJK characters break between any two glyphs, so a run of
// N of them wraps to ceil(N / charsPerLine) lines.

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

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

type DocRun = DocParagraph['runs'][number];

function para(opts: { text?: string; fontSize?: number; widowControl?: boolean } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
    widowControl: opts.widowControl,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A minimal anchored text-box shape (wps:wsp under wp:anchor). `wrapMode`
 *  defaults to topAndBottom so it reserves a full-width float band. */
function shapeRun(opts: {
  widthPt: number;
  heightPt: number;
  anchorYPt?: number;
  anchorYFromPara?: boolean;
  anchorXFromMargin?: boolean;
  wrapMode?: string | null;
  wrapSide?: string | null;
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
}): DocRun {
  const s: ShapeRun = {
    widthPt: opts.widthPt,
    heightPt: opts.heightPt,
    anchorXPt: 0,
    anchorYPt: opts.anchorYPt ?? 0,
    anchorXFromMargin: opts.anchorXFromMargin ?? true,
    anchorYFromPara: opts.anchorYFromPara ?? true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: opts.wrapMode === undefined ? 'topAndBottom' : opts.wrapMode,
    wrapSide: opts.wrapSide ?? null,
    distTop: opts.distTop ?? 0,
    distBottom: opts.distBottom ?? 0,
    distLeft: opts.distLeft ?? 0,
    distRight: opts.distRight ?? 0,
  };
  return { type: 'shape', ...s } as DocRun;
}

/** A paragraph carrying an explicit run list (e.g. an anchored shape, optionally
 *  followed by inline text). */
function paraWith(runs: DocRun[], opts: { fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

const sliceOf = (el: PaginatedBodyElement) =>
  (el as { lineSlice?: { start: number; end: number } }).lineSlice;

const colOf = (el: PaginatedBodyElement) => el.colIndex;

/** A body-level column break (ECMA-376 §17.3.1.20, hoisted by the parser). */
const colBreak = (): BodyElement => ({ type: 'columnBreak' } as BodyElement);

/** A body-level section break (ECMA-376 §17.6.x). `columns` is the ENDING
 *  section's `<w:cols>` (null/undefined ⇒ single full-width column). */
const sectionBreak = (
  kind: 'continuous' | 'nextPage' | 'oddPage' | 'evenPage',
  columns: SectionProps['columns'] = null,
): BodyElement => ({ type: 'sectionBreak', kind, columns } as BodyElement);

/** Text of a paragraph element (joins its text runs). */
const textOf = (el: PaginatedBodyElement): string =>
  el.type === 'paragraph'
    ? (el as unknown as DocParagraph).runs
        .filter((r) => r.type === 'text')
        .map((r) => (r as DocxTextRun).text)
        .join('')
    : '';

describe('computePages — empty-paragraph relocation (C2: §17.3.1.29)', () => {
  it('moves an unsplittable mark-only paragraph to the next page instead of overflowing the bottom margin', () => {
    // content height = 140 - 40 = 100; each empty mark = 20px → exactly 5 per page.
    const body = Array.from({ length: 7 }, () => para()); // 7 empty paragraphs
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(5); // page 1 fills exactly
    expect(pages[1].length).toBe(2); // overflow relocated, NOT clipped onto page 1
    // no page holds more than its 5-line capacity (would mean an overflow)
    for (const p of pages) expect(p.length).toBeLessThanOrEqual(5);
  });
});

describe('computePages — line-boundary splitting + widowControl (C1: §17.3.1.44)', () => {
  // contentW = 160, glyph advance = fontPx; at 20px → 8 chars/line. 48 chars → 6 lines.
  // content height = 100 → 5 lines (100px) fit per page.
  const sixLineText = 'あ'.repeat(48);

  it('avoids a widow: a single trailing line is not stranded on the next page (default widowControl on)', () => {
    const pages = computePages([para({ text: sixLineText })], section(), makeCtx());
    expect(pages.length).toBe(2);
    // Greedy fit is 5 lines on page 1; widowControl pulls one down so ≥2 carry over.
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 4 });
    expect(sliceOf(pages[1][0])).toEqual({ start: 4, end: 6 });
  });

  it('honors w:widowControl="off": the trailing single line is allowed (matches sample-9)', () => {
    const pages = computePages([para({ text: sixLineText, widowControl: false })], section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 }); // greedy 5 lines
    expect(sliceOf(pages[1][0])).toEqual({ start: 5, end: 6 }); // lone widow line allowed
  });
});

describe('computePages — anchored wrap-shape float exclusion (B: §20.4.2.16)', () => {
  // A topAndBottom anchored text-box SHAPE must reserve a float band exactly
  // like an anchored image does, so body text in following paragraphs flows
  // BELOW it instead of overlapping. Geometry: page content height =
  // 140 - 20 - 20 = 100pt. The shape is anchored at the first paragraph's top
  // (≈ marginTop = 20pt) with height 50 ⇒ band y∈[20,70]. The following text
  // paragraph is 2 lines × 20pt = 40pt.
  //
  // Without the shape registering a float (the bug): the shape paragraph's mark
  // (20pt: y20→40) + the 2-line text (40pt: y40→80) all fit ⇒ 1 page.
  // With the float (the fix): the mark flows below the band (y70→90) and the
  // text is pushed under the band, so its 2 lines (≥ y90 → ≈130) overflow the
  // 100pt content area ⇒ 2 pages.
  it('pushes following text below a topAndBottom shape (shape registers a float)', () => {
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'topAndBottom' })]),
      para({ text: 'あ'.repeat(16), fontSize: 20 }), // 160/20 = 8 chars/line → 2 lines
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // The text paragraph (with its float-displaced first line) is pushed to the
    // second page; page 1 holds only the shape-anchoring paragraph.
    const onPage2 = pages[1].some(
      (el) => el.type === 'paragraph' &&
        (el as unknown as DocParagraph).runs.some((r) => r.type === 'text'),
    );
    expect(onPage2).toBe(true);
  });

  it('does NOT reserve a band for a wrapNone shape (no float, text stays on one page)', () => {
    // Same geometry but wrapMode:'none' ⇒ no exclusion rect; everything fits on
    // one page. Guards against over-registering floats for non-wrapping shapes.
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'none' })]),
      para({ text: 'あ'.repeat(16), fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
  });
});

// ===== ECMA-376 §17.6.4 multi-column sections =====

describe('computeColumns — geometry (§17.6.4)', () => {
  it('count<=1 / no columns → one full-width column', () => {
    const cols = computeColumns(section());
    // contentW = 200 - 20 - 20 = 160; left = marginLeft = 20.
    expect(cols).toEqual([{ xPt: 20, wPt: 160 }]);
    // Explicit count:1 behaves the same.
    expect(computeColumns(section({ columns: { count: 1, spacePt: 36, equalWidth: true, sep: false, cols: [] } })))
      .toEqual([{ xPt: 20, wPt: 160 }]);
  });

  it('2 equal columns with a given space → correct x/w', () => {
    const cols = computeColumns(section({
      columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
    }));
    // colW = (160 - 1*20)/2 = 70. col0 at 20, col1 at 20 + 70 + 20 = 110.
    expect(cols).toEqual([
      { xPt: 20, wPt: 70 },
      { xPt: 110, wPt: 70 },
    ]);
  });

  it('3 equal columns tile the content band', () => {
    const cols = computeColumns(section({
      columns: { count: 3, spacePt: 10, equalWidth: true, sep: false, cols: [] },
    }));
    // colW = (160 - 2*10)/3 = 46.6667. Starts: 20, 76.6667, 133.3333.
    expect(cols).toHaveLength(3);
    expect(cols[0].xPt).toBeCloseTo(20, 6);
    expect(cols[0].wPt).toBeCloseTo(140 / 3, 6);
    expect(cols[1].xPt).toBeCloseTo(20 + 140 / 3 + 10, 6);
    expect(cols[2].xPt).toBeCloseTo(20 + 2 * (140 / 3 + 10), 6);
  });

  it('explicit <w:col> widths are used verbatim', () => {
    const cols = computeColumns(section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    }));
    // col0 at marginLeft=20, w=100; col1 at 20 + 100 + 12 = 132, w=48.
    expect(cols).toEqual([
      { xPt: 20, wPt: 100 },
      { xPt: 132, wPt: 48 },
    ]);
  });
});

describe('computePages — newspaper column flow (§17.6.4)', () => {
  const twoCol = (): SectionProps =>
    section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });

  // Geometry: page content height = 140 - 40 = 100; each single-line para = 20px
  // ⇒ 5 lines fill a column. colW = (160-20)/2 = 70.

  it('overflowing paragraphs flow into column 1 on the SAME page (pages.length unchanged)', () => {
    // 7 single-line paragraphs: 5 fill column 0, 2 spill into column 1 — still 1 page.
    const body = Array.from({ length: 7 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    expect(pages[0]).toHaveLength(7);
    // First 5 in column 0, last 2 in column 1.
    expect(pages[0].slice(0, 5).map(colOf)).toEqual([0, 0, 0, 0, 0]);
    expect(pages[0].slice(5).map(colOf)).toEqual([1, 1]);
  });

  it('fills both columns before starting a new page', () => {
    // 11 single-line paras: col0=5, col1=5, then the 11th starts page 2 col0.
    const body = Array.from({ length: 11 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(2);
    expect(pages[0]).toHaveLength(10);
    expect(pages[0].slice(0, 5).every((el) => colOf(el) === 0)).toBe(true);
    expect(pages[0].slice(5).every((el) => colOf(el) === 1)).toBe(true);
    expect(pages[1]).toHaveLength(1);
    expect(colOf(pages[1][0])).toBe(0); // new page resets to column 0
  });

  it('measures at the column width: a paragraph wraps to MORE lines than at full width', () => {
    // 6 CJK chars. Full width (160 → 8/line) = 1 line. Column width (70 → 3/line) = 2 lines.
    const sixChars = 'あ'.repeat(6);
    // Single-column baseline: one line, easily 1 page.
    const single = computePages([para({ text: sixChars, fontSize: 20 })], section(), makeCtx());
    expect(single.length).toBe(1);
    expect(sliceOf(single[0][0])).toBeUndefined(); // not split — fits on one line

    // Two-column: 70/20 = 3 chars/line ⇒ the same paragraph needs 2 lines. Fill
    // column 0 with 4 single-line paras (4 lines), so only 1 line is left in
    // column 0; the 2-line paragraph cannot fit there and is pushed to column 1.
    const body = [
      ...Array.from({ length: 4 }, (_, i) => para({ text: `p${i}`, fontSize: 20 })),
      para({ text: sixChars, fontSize: 20 }),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    // The 2-line paragraph (proof it wrapped to >1 line at column width) moved to
    // column 1 because only one line of room remained in column 0.
    const wrapped = pages[0].find((el) => textOf(el) === sixChars);
    expect(wrapped).toBeDefined();
    expect(colOf(wrapped as PaginatedBodyElement)).toBe(1);
  });

  it('a long paragraph splits across columns of the same page', () => {
    // 18 CJK chars at column width (3/line) = 6 lines. Column holds 5 lines, so it
    // splits: 5 lines in column 0, 1 line in column 1 (widowControl off to keep the
    // greedy split deterministic) — both on page 1.
    const body = [para({ text: 'あ'.repeat(18), fontSize: 20, widowControl: false })];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    expect(pages[0]).toHaveLength(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 });
    expect(colOf(pages[0][0])).toBe(0);
    expect(sliceOf(pages[0][1])).toEqual({ start: 5, end: 6 });
    expect(colOf(pages[0][1])).toBe(1);
  });
});

describe('computePages — explicit column break (§17.3.1.20)', () => {
  const twoCol = (): SectionProps =>
    section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });

  it('a column break forces the next paragraph into column 1 (same page)', () => {
    const body = [
      para({ text: 'a', fontSize: 20 }),
      colBreak(),
      para({ text: 'b', fontSize: 20 }),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    // The break is consumed (not emitted as a body element); two paragraphs remain.
    const paras = pages[0].filter((el) => el.type === 'paragraph');
    expect(paras).toHaveLength(2);
    expect(colOf(paras[0])).toBe(0);
    expect(colOf(paras[1])).toBe(1);
  });

  it('a column break in the LAST column moves to the next page (column 0)', () => {
    const body = [
      para({ text: 'a', fontSize: 20 }),
      colBreak(), // → column 1
      para({ text: 'b', fontSize: 20 }),
      colBreak(), // last column → page 2, column 0
      para({ text: 'c', fontSize: 20 }),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(2);
    expect(colOf(pages[0].filter((e) => e.type === 'paragraph')[0])).toBe(0); // a
    expect(colOf(pages[0].filter((e) => e.type === 'paragraph')[1])).toBe(1); // b
    expect(textOf(pages[1][0])).toBe('c');
    expect(colOf(pages[1][0])).toBe(0); // c on page 2, column 0
  });
});

describe('computePages — single-column regression (§17.6.4 absent)', () => {
  it('a single-column section paginates identically (no colIndex > 0, full width)', () => {
    // Reuse the empty-paragraph relocation scenario: 7 empty paras, 5/page.
    const body = Array.from({ length: 7 }, () => para());
    const withCols = computePages(body, section(), makeCtx());
    expect(withCols.length).toBe(2);
    expect(withCols[0].length).toBe(5);
    expect(withCols[1].length).toBe(2);
    // Every placed element is column 0 (or untagged) — never a higher column.
    for (const page of withCols) {
      for (const el of page) expect(colOf(el) ?? 0).toBe(0);
    }
  });

  it('a wrapping paragraph splits the same way single-column as before', () => {
    // 48 CJK chars at full width (8/line) = 6 lines; 5 fit per page; widowControl
    // off ⇒ greedy 5 + 1. Identical to the pre-column behavior (guards no regression).
    const pages = computePages(
      [para({ text: 'あ'.repeat(48), fontSize: 20, widowControl: false })],
      section(),
      makeCtx(),
    );
    expect(pages.length).toBe(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 });
    expect(sliceOf(pages[1][0])).toEqual({ start: 5, end: 6 });
    expect(colOf(pages[0][0]) ?? 0).toBe(0);
    expect(colOf(pages[1][0]) ?? 0).toBe(0);
  });
});

describe('computePages — per-section columns (§17.6.4, regression)', () => {
  // The sample-5 shape: section 1 is single-column (no <w:cols num>), only the
  // FINAL section is 2-column. Section 1's content MUST lay out in ONE full-width
  // column — not in the final section's 2-column grid. `section.columns` carries
  // the FINAL (body-level) section's columns; each mid-body section's columns ride
  // on its SectionBreak marker (here None ⇒ single column).
  const colGeomOf = (el: PaginatedBodyElement): { xPt: number; wPt: number }[] | undefined =>
    (el as { colGeom?: { xPt: number; wPt: number }[] }).colGeom;

  const finalTwoCol = (): SectionProps =>
    section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });

  it("section 1 lays out in ONE full-width column even though the FINAL section is 2-col", () => {
    // 8 CJK chars: full width (160 → 8/line) = 1 line; the buggy 2-col path (70px →
    // 3/line) would wrap it to 3 lines. Section 1 = this para, ended by a nextPage
    // SectionBreak (cols None). Final section = one 2-col para.
    const body: BodyElement[] = [
      para({ text: 'あ'.repeat(8), fontSize: 20 }),
      sectionBreak('nextPage', null),
      para({ text: 'final', fontSize: 20 }),
    ];
    const pages = computePages(body, finalTwoCol(), makeCtx());
    // SectionBreak is consumed (not emitted); section 1 on page 1, final on page 2.
    expect(pages.length).toBe(2);
    const sec1Para = pages[0][0];
    // Full-width single column ⇒ NOT split (still one line) and colGeom length 1.
    expect(sliceOf(sec1Para)).toBeUndefined();
    expect(colOf(sec1Para) ?? 0).toBe(0);
    expect(colGeomOf(sec1Para)).toEqual([{ xPt: 20, wPt: 160 }]);
    // The FINAL section's paragraph gets the 2-column geometry.
    const finalPara = pages[1][0];
    expect(colGeomOf(finalPara)).toEqual([
      { xPt: 20, wPt: 70 },
      { xPt: 110, wPt: 70 },
    ]);
  });

  it('section 1 fills the full column height (5 lines/page) before paginating, not the half-width 2-col height', () => {
    // 5 single-line paragraphs fill a full-width column exactly (content height 100
    // / 20px = 5). All land on page 1, column 0, full-width geometry. Under the bug
    // they would have been squeezed into 2-col width and the flow would differ.
    const body: BodyElement[] = [
      ...Array.from({ length: 5 }, (_, i) => para({ text: `s${i}`, fontSize: 20 })),
      sectionBreak('nextPage', null),
      para({ text: 'final', fontSize: 20 }),
    ];
    const pages = computePages(body, finalTwoCol(), makeCtx());
    expect(pages[0]).toHaveLength(5);
    for (const el of pages[0]) {
      expect(colOf(el) ?? 0).toBe(0);
      expect(colGeomOf(el)).toEqual([{ xPt: 20, wPt: 160 }]);
    }
  });

  it('a continuous section break keeps both single-column sections on the SAME page, stacked', () => {
    // ECMA-376 §17.6.22: a section's `<w:type>` describes how THAT section begins
    // relative to the previous one, and lives on the sectPr that ENDS the section
    // (the NEXT marker in document order). So the A→B break is governed by section
    // B's type — the kind of the marker that ENDS B (idx 3) — and the B→final break
    // by the final (body) section's start type. Here B is continuous (stacks below
    // A on page 1) and the final 2-col section is nextPage (default) → page 2. Both
    // A and B are single-column full width and must NOT reset to the page top.
    const body: BodyElement[] = [
      para({ text: 'A', fontSize: 20 }),
      sectionBreak('continuous', null), // ends section A (A's own start type)
      para({ text: 'B', fontSize: 20 }),
      sectionBreak('continuous', null), // ends section B ⇒ B begins continuously (A→B same page)
      para({ text: 'final', fontSize: 20 }),
    ];
    const pages = computePages(body, finalTwoCol(), makeCtx());
    expect(pages.length).toBe(2);
    // A and B share page 1 (B continuous), final on page 2 (final section nextPage).
    expect(pages[0].map(textOf)).toEqual(['A', 'B']);
    expect(pages[1].map(textOf)).toEqual(['final']);
    for (const el of pages[0]) {
      expect(colGeomOf(el)).toEqual([{ xPt: 20, wPt: 160 }]);
    }
  });

  it('the break INTO a section is governed by that section start type, not the prior marker (§17.6.22)', () => {
    // Regression for the off-by-one: a title section whose own sectPr is nextPage
    // (the spec default) followed by a continuous body section must NOT page-break —
    // the boundary is governed by the FOLLOWING (body) section's start type. Mirrors
    // sample-13, where Word keeps the 2-col body on the title page.
    const body: BodyElement[] = [
      para({ text: 'title', fontSize: 20 }),
      sectionBreak('nextPage', null), // ends the title section (its own start type)
      para({ text: 'body', fontSize: 20 }),
    ];
    const sec = { ...finalTwoCol(), columns: null, sectionStart: 'continuous' };
    const pages = computePages(body, sec, makeCtx());
    expect(pages.length).toBe(1);
    expect(pages[0].map(textOf)).toEqual(['title', 'body']);
  });

  it('single-section 2-col docs are UNCHANGED (no SectionBreak markers ⇒ whole body uses section.columns)', () => {
    // Guard: with NO section breaks, columns = section.columns for the whole body,
    // identical to the pre-fix global-columns behavior (sample-10's case).
    const body = Array.from({ length: 7 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, finalTwoCol(), makeCtx());
    // 5 fill col0, 2 spill into col1 — one page (the existing newspaper-flow test).
    expect(pages.length).toBe(1);
    expect(pages[0].slice(0, 5).map(colOf)).toEqual([0, 0, 0, 0, 0]);
    expect(pages[0].slice(5).map(colOf)).toEqual([1, 1]);
    // Every element carries the same 2-col geometry (the body-level section's).
    for (const el of pages[0]) {
      expect(colGeomOf(el)).toEqual([
        { xPt: 20, wPt: 70 },
        { xPt: 110, wPt: 70 },
      ]);
    }
  });
});
