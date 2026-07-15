import { describe, expect, it, vi } from 'vitest';
import { layoutDocument } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, LineNumbering, SectionProps } from './types.js';
import type { ParagraphLayout } from './layout/types.js';
import type { PlacedFragment } from './layout-fragments.js';

const FONT = 'Synthetic Untabled Serif';

function measurementContext(): OffscreenCanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText(text: string) {
      const size = Number.parseFloat(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} } as CanvasGradient; },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
    globalAlpha: 1, lineCap: 'butt', lineJoin: 'miter',
  } as unknown as OffscreenCanvasRenderingContext2D;
}

vi.stubGlobal('OffscreenCanvas', class {
  constructor(_width: number, _height: number) {}
  getContext() { return measurementContext(); }
});

interface PrivateSectionPlacementWire {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: LineNumbering | null;
}

type InternalSectionBreak = Extract<BodyElement, { type: 'sectionBreak' }> & {
  readonly __sectionPlacement: PrivateSectionPlacementWire;
};

function para(text: string, overrides: Partial<DocParagraph> = {}): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: FONT,
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
      hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: FONT, widowControl: false,
    ...overrides,
  } as unknown as BodyElement;
}

function marker(kind: string, placement: PrivateSectionPlacementWire): BodyElement {
  return {
    type: 'sectionBreak', kind, columns: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
    __sectionPlacement: placement,
  } as InternalSectionBreak;
}

function section(over: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 200,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 40,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
    ...over,
  } as SectionProps;
}

function doc(body: BodyElement[], finalSection: SectionProps): DocxDocumentModel {
  return {
    section: finalSection,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

function fragments(model: DocxDocumentModel): PlacedFragment[][] {
  return (layoutDocument(model) as unknown as { pages: Array<{ fragments: PlacedFragment[] }> })
    .pages.map((page) => page.fragments);
}

function paragraphFragments(page: readonly PlacedFragment[]): Array<PlacedFragment & { fragment: ParagraphLayout }> {
  return page.filter((placed): placed is PlacedFragment & { fragment: ParagraphLayout } =>
    placed.fragment.kind === 'paragraph');
}

describe('retained per-section flow placement', () => {
  it('uses the non-final section vAlign and numbering on its own next-page region', () => {
    const firstLn: LineNumbering = { countBy: 1, start: 4, distance: 10, restart: 'newSection' };
    const pages = fragments(doc([
      para('FIRST'),
      marker('nextPage', { sectionId: 'section:0', vAlign: 'center', lineNumbering: firstLn }),
      para('FINAL'),
    ], section({ sectionStart: 'nextPage', vAlign: 'top', lineNumbering: null })));

    expect(pages).toHaveLength(2);
    const first = paragraphFragments(pages[0]!)[0]!;
    const final = paragraphFragments(pages[1]!)[0]!;
    expect(first.yPt).toBeGreaterThan(final.yPt + 50);
    expect(first.fragment.lineNumbers?.map((line) => line.counterValue)).toEqual([4]);
    expect(first.fragment.lineNumbers?.[0]?.paintOps).toHaveLength(1);
    expect(final.fragment.lineNumbers).toBeUndefined();
  });

  it('keeps continuous mixed-section vAlign regions independent on one page', () => {
    const page = fragments(doc([
      para('FIRST'),
      marker('nextPage', { sectionId: 'section:0', vAlign: 'top', lineNumbering: null }),
      para('FINAL'),
    ], section({ sectionStart: 'continuous', vAlign: 'bottom', lineNumbering: null })))[0]!;
    const [first, final] = paragraphFragments(page);

    expect(first?.yPt).toBeCloseTo(20, 1);
    expect(final?.yPt).toBeGreaterThan(150);
  });

  it('restart=newSection resets exact counters and retained paint geometry at a continuous boundary', () => {
    const page = fragments(doc([
      para('a'), para('b'),
      marker('nextPage', {
        sectionId: 'section:0', vAlign: 'top',
        lineNumbering: { countBy: 2, start: 3, distance: 10, restart: 'newSection' },
      }),
      para('c'), para('d'),
    ], section({
      sectionStart: 'continuous', vAlign: 'top',
      lineNumbering: { countBy: 2, start: 7, distance: 10, restart: 'newSection' },
    })))[0]!;
    const layouts = paragraphFragments(page).map((placed) => placed.fragment);

    expect(layouts.flatMap((layout) => layout.lineNumbers?.map((line) => line.counterValue) ?? []))
      .toEqual([3, 4, 7, 8]);
    expect(layouts.flatMap((layout) => layout.lineNumbers?.map((line) => line.paintOps.length) ?? []))
      .toEqual([0, 1, 0, 1]);
    for (const lineNumber of layouts.flatMap((layout) => layout.lineNumbers ?? [])) {
      expect(Number.isFinite(lineNumber.bounds.xPt)).toBe(true);
      expect(Number.isFinite(lineNumber.bounds.yPt)).toBe(true);
      expect(lineNumber.bounds.widthPt).toBeGreaterThan(0);
      if (lineNumber.paintOps[0]) {
        expect(lineNumber.paintOps[0].origin.xPt)
          .toBeCloseTo(lineNumber.bounds.xPt + lineNumber.bounds.widthPt, 6);
      }
    }
  });

  it('excludes frame paragraphs from body height and line numbering', () => {
    const page = paragraphFragments(fragments(doc([
      para('FRAME', {
        framePr: {
          dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
          hRule: 'auto', hSpace: 0, vSpace: 0, w: 40,
        },
      }),
      para('BODY'),
    ], section({
      vAlign: 'bottom',
      lineNumbering: { countBy: 1, start: 1, distance: 10, restart: 'newPage' },
    })))[0]!);
    const frame = page.find(({ fragment }) => !fragment.ordinaryFlow)!;
    const body = page.find(({ fragment }) => fragment.ordinaryFlow)!;

    expect(frame.fragment.lineNumbers).toBeUndefined();
    expect(body.fragment.lineNumbers?.map((line) => line.counterValue)).toEqual([1]);
    expect(frame.heightPt).toBe(0);
    // A text-anchored frame follows the same retained vAlign translation as its
    // host flow; its own height does not enlarge that alignment calculation.
    expect(frame.yPt).toBeGreaterThan(100);
    expect(body.yPt).toBeGreaterThan(100);
  });

  it('corrects retained frame column ownership after pagination moves to column two', () => {
    const body = [
      ...Array.from({ length: 7 }, (_, index) => para(`flow-${index}`)),
      para('FRAME', {
        framePr: {
          dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
          hRule: 'auto', hSpace: 0, vSpace: 0, w: 30,
        },
      }),
      para('anchor'),
    ];
    const pages = fragments(doc(body, section({
      pageHeight: 100, marginTop: 10, marginBottom: 10,
      columns: { count: 2, spacePt: 10, equalWidth: true, sep: false, cols: [] },
    })));
    const frame = pages.flat().find((placed) =>
      placed.fragment.kind === 'paragraph' && !placed.fragment.ordinaryFlow);
    expect(frame?.columnIndex).toBe(1);
  });

  it('does not invent a vAlign translation for a frame-only region', () => {
    const page = paragraphFragments(fragments(doc([
      para('FRAME', {
        framePr: {
          dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
          hRule: 'auto', hSpace: 0, vSpace: 0, w: 40,
        },
      }),
    ], section({ vAlign: 'bottom' })))[0]!);
    const frame = page.find(({ fragment }) => !fragment.ordinaryFlow)!;

    expect(frame.heightPt).toBe(0);
    expect(frame.yPt).toBeLessThan(50);
  });

  it.each([
    { anchor: 'page', expectedY: 5 },
    { anchor: 'margin', expectedY: 25 },
  ])('keeps a $anchor-anchored frame absolute across bottom vAlign and header reserve', ({ anchor, expectedY }) => {
    const makeModel = (header: boolean) => {
      const model = doc([
        para('FRAME', {
          framePr: {
            dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: anchor,
            hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, y: 5,
          },
        }),
        para('BODY'),
      ], section({ vAlign: 'bottom', headerDistance: 0 }));
      if (header) {
        model.headers.default = {
          body: Array.from({ length: 4 }, (_, index) => para(`HEADER-${index}`)),
        };
      }
      return model;
    };
    const frameY = (model: DocxDocumentModel) => paragraphFragments(fragments(model)[0]!)
      .find(({ fragment }) => !fragment.ordinaryFlow)!.yPt;

    expect(frameY(makeModel(false))).toBeCloseTo(expectedY, 6);
    expect(frameY(makeModel(true))).toBeCloseTo(expectedY, 6);
  });
});
