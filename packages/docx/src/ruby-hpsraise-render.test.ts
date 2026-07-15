import { describe, expect, it } from 'vitest';
import { buildSegments, rubyAscentReservePx } from './line-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

interface GlyphCall {
  text: string;
  x: number;
  y: number;
  font: string;
}

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  glyphs: GlyphCall[];
} {
  let font = '10px serif';
  let letterSpacing = '0px';
  const glyphs: GlyphCall[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) { letterSpacing = value; },
    measureText(text: string) {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) { glyphs.push({ text, x, y, font }); },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection, globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, glyphs };
}

function textRun(hpsRaisePt?: number): DocxTextRun {
  return {
    text: '漢', bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
    ruby: {
      text: 'かん',
      fontSizePt: 8,
      ...(hpsRaisePt != null ? { hpsRaisePt } : {}),
    },
  };
}

function documentWithRuby(hpsRaisePt?: number): DocxDocumentModel {
  const paragraph: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(hpsRaisePt) }],
    defaultFontSize: 12, defaultFontFamily: 'NotInMetrics', widowControl: false,
  };
  const section: SectionProps = {
    pageWidth: 200, pageHeight: 200,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  };
  return {
    section,
    body: [{ type: 'paragraph', ...paragraph } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function bodyRubyBaselineDistance(hpsRaisePt?: number, scale = 2): Promise<number> {
  const { canvas, glyphs } = makeRecordingCanvas();
  await renderDocumentToCanvas(documentWithRuby(hpsRaisePt), canvas, 0, {
    dpr: 1,
    width: 200 * scale,
  });
  const base = glyphs.find((glyph) => glyph.text === '漢');
  const ruby = glyphs.find((glyph) => /[かん]/.test(glyph.text));
  expect(base, 'base glyph drawn').toBeDefined();
  expect(ruby, 'ruby glyph drawn').toBeDefined();
  return base!.y - ruby!.y;
}

describe('§17.3.3.12 w:hpsRaise ruby geometry', () => {
  it('draws horizontal body ruby hpsRaise above the base baseline', async () => {
    expect(await bodyRubyBaselineDistance(14)).toBeCloseTo(14, 8);
  });

  it('touches retained base and guide ink when hpsRaise is absent', async () => {
    // The injected authoritative ink is 9.6pt above the base baseline and
    // 1.6pt below the 8pt guide baseline. Point-space geometry is scale-free.
    expect(await bodyRubyBaselineDistance(undefined, 2)).toBeCloseTo(11.2, 8);
  });

  it('treats explicit hpsRaise zero as a zero reservation and zero draw offset', async () => {
    expect(rubyAscentReservePx(8, 0, 2)).toBe(0);
    expect(await bodyRubyBaselineDistance(0)).toBe(0);
  });

  it('keeps absent hpsRaise off the emitted segment own-property set', () => {
    const environment = {
      defaultFontSize: 12,
      defaultFontFamily: 'NotInMetrics',
      fontFamilyClasses: {},
      pageIndex: 0,
      totalPages: 1,
    };
    const [absent] = buildSegments(
      [{ type: 'text', ...textRun() }],
      environment,
    );
    const [zero] = buildSegments(
      [{ type: 'text', ...textRun(0) }],
      environment,
    );
    expect(absent && 'ruby' in absent && absent.ruby).toBeDefined();
    expect(zero && 'ruby' in zero && zero.ruby).toBeDefined();
    if (!absent || !('ruby' in absent) || !zero || !('ruby' in zero)) {
      throw new Error('expected text segments');
    }
    expect(Object.hasOwn(absent.ruby!, 'hpsRaisePt')).toBe(false);
    expect(Object.hasOwn(zero.ruby!, 'hpsRaisePt')).toBe(true);
    expect(zero.ruby!.hpsRaisePt).toBe(0);
  });
});
