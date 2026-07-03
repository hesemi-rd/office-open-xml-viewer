import { describe, it, expect } from 'vitest';
import { computePages } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  FramePr,
  SectionProps,
  PaginatedBodyElement,
} from './types';

// Unit tests for the keep-with-anchor pagination of a page-overflowing text
// frame (ECMA-376 §17.3.1.11 `<w:framePr>`). §17.3.1.11 pins only the frame's
// size/position; keeping an undivided frame with its anchor context on the next
// page is Word runtime behaviour, the same "keep on page" semantics as a
// paragraph-anchored image float. These assertions guard that behaviour, which
// the private-sample VRT (drop caps only, no page-boundary frame) cannot cover.
//
// The stub canvas mirrors pagination.test.ts: glyph advance = charCount × fontPx
// and the font box = 0.8/0.2 em, so a single line is exactly fontPx tall.

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

// pageWidth 200 / pageHeight 140, margins 20 ⇒ content band 160×100 (bodyTop 20).
function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
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

// Full FramePr with the spec defaults; callers override only the axis under test.
function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none',
    lines: 1,
    wrap: 'around',
    hAnchor: 'text',
    vAnchor: 'text',
    hRule: 'auto',
    hSpace: 0,
    vSpace: 0,
    ...over,
  };
}

function para(opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A frame paragraph (`w:framePr`) carrying a single text run. */
function framePara(fp: FramePr, opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [textRun(opts.text ?? 'F', fontSize)],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
    framePr: fp,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** True when a page holds the frame paragraph (identified by its framePr). */
const hasFrame = (page: PaginatedBodyElement[]): boolean =>
  page.some((el) => el.type === 'paragraph' && (el as unknown as DocParagraph).framePr != null);

/** Text of a paragraph element (joins its text runs). */
const textOf = (el: PaginatedBodyElement): string =>
  el.type === 'paragraph'
    ? (el as unknown as DocParagraph).runs
        .filter((r) => r.type === 'text')
        .map((r) => (r as DocxTextRun).text)
        .join('')
    : '';

/** True when a page holds the anchor paragraph (matched by its text). */
const hasAnchorText = (page: PaginatedBodyElement[], text: string): boolean =>
  page.some((el) => textOf(el) === text);

/** The newspaper column an element landed in. */
const colOf = (el: PaginatedBodyElement): number | undefined => el.colIndex;

/** Find the (first) frame paragraph element on a page. */
const frameEl = (page: PaginatedBodyElement[]): PaginatedBodyElement | undefined =>
  page.find((el) => el.type === 'paragraph' && (el as unknown as DocParagraph).framePr != null);

describe('computePages — text-frame keep-with-anchor (§17.3.1.11)', () => {
  // Content band 160×100, bodyTop 20. Frame: vAnchor="text", hRule="exact",
  // h=60 ⇒ frame body box [paraTop, paraTop+60]. With N leading 20pt lines the
  // frame paragraph's in-flow top is at y=20N; the frame overflows the 100pt
  // content area once 20N + 60 > 100 ⇒ N ≥ 3 (y=60).

  it('relocates a text-anchored frame + its anchor text to the next page when it overflows the bottom', () => {
    // 3 leading lines (y advances 20→40→60), then the frame (needs 60pt below
    // its top: 60+60 > 100 ⇒ overflow), then the anchor text paragraph.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      framePara(frame({ vAnchor: 'text', hRule: 'exact', h: 60 }), { text: 'F' }),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // The frame does NOT stay on page 1 (would overflow); it moves to page 2…
    expect(hasFrame(pages[0])).toBe(false);
    expect(hasFrame(pages[1])).toBe(true);
    // …and its trailing anchor text follows it onto page 2 (kept together).
    expect(hasAnchorText(pages[1], 'anchor')).toBe(true);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(false);
    // Page 1 keeps only the three leading lines (no float band leaked over).
    expect(pages[0].map(textOf)).toEqual(['a', 'b', 'c']);
  });

  it('relocates an overflowing frame to the NEXT COLUMN (not a new page) in a multi-column section', () => {
    // 2 equal columns: colW = (160-20)/2 = 70; content height 100 (5 × 20pt per
    // column). Column 0 gets 3 leading lines (y 20→40→60); the frame body box
    // (h=60) then overflows column 0's bottom (60+60 > 100). With a column still
    // available on the page, it relocates to column 1 — NOT a new page.
    const twoCol = section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      framePara(frame({ vAnchor: 'text', hAnchor: 'text', hRule: 'exact', h: 60 }), { text: 'F' }),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, twoCol, makeCtx());
    // Still a single page (column 1 absorbed the frame + anchor).
    expect(pages.length).toBe(1);
    const f = frameEl(pages[0]);
    expect(f).toBeDefined();
    expect(colOf(f as PaginatedBodyElement)).toBe(1); // moved into column 1
    // The three leading lines (a/b/c) stayed in column 0.
    const leadCols = pages[0]
      .filter((el) => ['a', 'b', 'c'].includes(textOf(el)))
      .map(colOf);
    expect(leadCols).toEqual([0, 0, 0]);
    // The trailing anchor text follows the frame into column 1.
    const anchor = pages[0].find((el) => textOf(el) === 'anchor');
    expect(anchor).toBeDefined();
    expect(colOf(anchor as PaginatedBodyElement)).toBe(1);
  });

  it('keeps a text-anchored frame in place when it fits (near the top of the page)', () => {
    // Only 1 leading line (y=20). Frame body box [20,80] fits within [0,100] ⇒
    // no relocation. Everything stays on page 1.
    const body = [
      para({ text: 'a' }),
      framePara(frame({ vAnchor: 'text', hRule: 'exact', h: 60 }), { text: 'F' }),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFrame(pages[0])).toBe(true);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
  });

  it('does NOT relocate (or loop) a frame taller than the page content area', () => {
    // h=150 > content height 100: the frame can never fit on any page, so
    // relocating would loop forever. It is left in place and allowed to overflow.
    // The real assertion is that this terminates (no timeout / infinite paging).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      framePara(frame({ vAnchor: 'text', hRule: 'exact', h: 150 }), { text: 'F' }),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    // Frame stays on page 1 with its three leading lines (no relocation).
    expect(hasFrame(pages[0])).toBe(true);
    // A frame that adds no flow height leaves the anchor on the same page.
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
    expect(pages.length).toBe(1);
  });

  it('does NOT relocate a page-anchored frame that overflows (absolute y ⇒ kept on page, box clamped)', () => {
    // vAnchor="page", y=90: the frame is pinned at page-y 90 with h=60 ⇒ its
    // requested bottom is 150, past the 140 page edge. An absolute page position is
    // the SAME on any page, so PAGINATION cannot help — the frame stays on page 1
    // (no relocation). Word does not let it overflow: computeFrameBox clamps the box
    // UP into the page (to pageH − h; see frame-geometry.test.ts "clamp"). This test
    // guards only the pagination decision (page count / no relocation); the clamped
    // box y is asserted in the geometry suite.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      framePara(frame({ vAnchor: 'page', hRule: 'exact', h: 60, y: 90 }), { text: 'F' }),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFrame(pages[0])).toBe(true);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
  });

  it('does NOT relocate a margin-anchored frame that overflows (absolute y ⇒ kept on page, box clamped)', () => {
    // vAnchor="margin", y=70: frame at margin-top(20)+70 = 90, h=60 ⇒ requested
    // bottom 150, past the bottom margin. Absolute ⇒ kept on page 1 (mirrors
    // vAnchor="page"); the box is clamped up into the margin band by computeFrameBox
    // (geometry suite). This test guards only the pagination decision.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      framePara(frame({ vAnchor: 'margin', hRule: 'exact', h: 60, y: 70 }), { text: 'F' }),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFrame(pages[0])).toBe(true);
  });

  it('leaves a drop-cap frame (sample-11 shape) unchanged — small frame, no overflow', () => {
    // A dropCap="drop" frame sized by `lines` (2 × the 20pt anchor line = 40pt)
    // near the top of the page never overflows, so pagination is unchanged: the
    // drop-cap frame and its anchor paragraph sit together on page 1.
    const body = [
      framePara(
        frame({ dropCap: 'drop', lines: 2, wrap: 'around', vAnchor: 'text', hAnchor: 'text' }),
        { text: 'D', fontSize: 40 },
      ),
      para({ text: 'あ'.repeat(8), fontSize: 20 }), // anchor body text
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFrame(pages[0])).toBe(true);
    // The drop cap registers a wrap float on page 1 (following text wraps around
    // it), confirming the register path still runs when no relocation happens.
    expect(pages[0].length).toBe(2);
  });
});
