import { describe, expect, it } from 'vitest';
import { buildSegments, layoutLines, rescaleLayoutLines, type LineLayoutEnvironment } from '../line-layout.js';
import { createLayoutServices } from '../renderer.js';
import type { DocRun, DocxDocumentModel } from '../types.js';
import type { TextLayoutService } from './text.js';
import { mathAstResourceKey } from './resources.js';
import { privateResourceLookupOf } from './runtime-state.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 8,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function model(overrides: Partial<DocxDocumentModel> = {}): DocxDocumentModel {
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    ...overrides,
  };
}

function textRun(text: string, extra: Record<string, unknown> = {}): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Authored Sans',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  } as DocRun;
}

describe('production layout service integration', () => {
  it('routes every normal run segmentation and measurement through the injected text service', () => {
    const base = createLayoutServices(model(), { measureContext: measureContext() });
    let calls = 0;
    const countingText: TextLayoutService = Object.freeze({
      ...base.text,
      shape(request: Parameters<TextLayoutService['shape']>[0]) {
        calls += 1;
        return base.text.shape(request);
      },
    });
    const services = Object.freeze({ ...base, text: countingText });
    const environment: LineLayoutEnvironment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const segments = buildSegments([textRun('first'), textRun('second')], environment);
    const afterSegmentation = calls;
    const lines = layoutLines(measureContext(), segments, 300, 0, 1);

    expect(afterSegmentation).toBeGreaterThanOrEqual(2);
    expect(calls).toBeGreaterThan(afterSegmentation);
    const afterLayout = calls;
    rescaleLayoutLines(lines, 2, measureContext(), {}, 0);
    expect(calls).toBeGreaterThan(afterLayout);
  });

  it('carries the w:kern threshold through the text service measure adapter', () => {
    let fontKerning: CanvasFontKerning = 'auto';
    const states: CanvasFontKerning[] = [];
    const ctx = {
      font: '',
      letterSpacing: '0px',
      get fontKerning() { return fontKerning; },
      set fontKerning(value: CanvasFontKerning) { fontKerning = value; },
      measureText(text: string) {
        states.push(fontKerning);
        const width = fontKerning === 'normal' ? 40 : fontKerning === 'none' ? 30 : 20;
        return {
          width: text ? width : 0,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const services = createLayoutServices(model(), { measureContext: ctx });
    const environment: LineLayoutEnvironment = { pageIndex: 0, totalPages: 1, layoutServices: services };
    const measure = (kerning: number) => {
      const segments = buildSegments([textRun('AV', { fontSize: 10, kerning })], environment);
      return layoutLines(ctx, segments, 300, 0, 1)[0].segments[0].measuredWidth;
    };

    expect(measure(5)).toBe(40);
    expect(measure(20)).toBe(30);
    expect(states).toContain('normal');
    expect(states).toContain('none');
    expect(ctx.fontKerning).toBe('auto');
  });

  it('keeps hint-protected eastAsia text on non-cs formatting inside an rtl run', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const run = textRun('A国', {
      fontFamily: 'Latin Face',
      fontFamilyEastAsia: 'EA Face',
      fontFamilyCs: 'CS Face',
      fontHint: 'eastAsia',
      langEastAsia: 'zh-cn',
      rtl: true,
      cs: true,
      fontSize: 10,
      fontSizeCs: 20,
      bold: false,
      boldCs: true,
    });
    const segments = buildSegments([run], { pageIndex: 0, totalPages: 1, layoutServices: services });
    const textSegments = segments.filter((segment) => 'text' in segment && segment.text.length > 0);

    expect(textSegments.map((segment) => 'text' in segment ? {
      text: segment.text,
      fontFamily: segment.fontFamily,
      fontSize: segment.fontSize,
      bold: segment.bold,
    } : null)).toEqual([
      { text: 'A', fontFamily: 'sans-serif', fontSize: 20, bold: true },
      { text: '国', fontFamily: 'sans-serif', fontSize: 10, bold: false },
    ]);
    expect(textSegments.map((segment) => 'text' in segment
      ? segment.textShapeRequest?.fonts.complexScript === 'CS Face' && segment.textShapeRequest?.complexScript
      : false)).toEqual([true, false]);
  });

  it('keeps a single pure-CJK hint-protected rtl span on non-cs formatting', () => {
    const services = createLayoutServices(model(), { measureContext: measureContext() });
    const [segment] = buildSegments([textRun('国', {
      fontFamily: 'Latin Face', fontFamilyEastAsia: 'EA Face', fontFamilyCs: 'CS Face',
      fontHint: 'eastAsia', langEastAsia: 'zh-cn', rtl: true, cs: true,
      fontSize: 10, fontSizeCs: 20, bold: false, boldCs: true,
    })], { pageIndex: 0, totalPages: 1, layoutServices: services });

    expect(segment).toMatchObject({ text: '国', fontSize: 10, bold: false, fontFamily: 'sans-serif' });
    expect('text' in segment && segment.textShapeRequest?.complexScript).toBe(false);
  });

  it('plumbs the selected eastAsia font charset from fontTable into slot selection', () => {
    const doc = model() as DocxDocumentModel & { fontFamilyCharsets: Record<string, string> };
    doc.fontFamilyCharsets = { 'EA Face': '86' };
    const services = createLayoutServices(doc, { measureContext: measureContext() });
    const shaped = services.text.shape({
      text: '\u0100',
      fontSizePt: 10,
      fontHint: 'eastAsia',
      eastAsiaLanguage: 'ja-jp',
      fonts: { ascii: 'Latin Face', highAnsi: 'Latin Face', eastAsia: 'EA Face' },
    });

    expect(shaped.spans[0]?.script).toBe('eastAsia');
  });

  it('makes local availability, geometry, diagnostics, and fingerprints truthful and worker-stable', () => {
    const fonts: string[] = [];
    const ctx = {
      ...measureContext(),
      get font() { return fonts.at(-1) ?? ''; },
      set font(value: string) { fonts.push(value); },
      measureText(text: string) {
        const width = (fonts.at(-1) ?? '').includes('__local_times_bi') ? 17 : 5;
        return {
          width: text.length * width,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as CanvasRenderingContext2D;
    const localMetrics = {
      'times new roman:700:italic': {
        family: '__local_times_bi', lineHeightRatio: 1.2,
        requestedFamily: 'Times New Roman', weight: 700, style: 'italic' as const,
      },
    };
    const shape = (services: ReturnType<typeof createLayoutServices>) => services.text.shape({
      text: 'AV', fontSizePt: 10, weight: 700, style: 'italic',
      fonts: { ascii: 'Times New Roman' },
    });
    const present = createLayoutServices(model(), { measureContext: ctx, localMetrics });
    const worker = createLayoutServices(model(), { measureContext: ctx, localMetrics });
    const absent = createLayoutServices(model(), { measureContext: ctx, localMetrics: {} });

    expect(shape(present).spans[0]?.font).toMatchObject({ source: 'local', resolvedFamily: '__local_times_bi' });
    expect(shape(present).advancePt).toBe(34);
    expect(shape(absent).spans[0]?.font).toMatchObject({ source: 'generic', resolvedFamily: 'sans-serif' });
    expect(shape(absent).advancePt).toBe(10);
    expect(present.text.fingerprint).toBe(worker.text.fingerprint);
    expect(present.text.fingerprint).not.toBe(absent.text.fingerprint);
  });

  it('inventories only successfully registered faces and labels Office replacements as substitutions', () => {
    const doc = model({
      majorFont: 'Calibri',
      embeddedFonts: [{ fontName: 'Broken Embedded', partPath: 'word/fonts/missing.odttf', fontKey: '', style: 'regular' }],
    });
    const failed = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [],
      googleFaces: [],
      useGoogleFonts: true,
    });
    const missingEmbedded = failed.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Broken Embedded' } });
    const missingGoogle = failed.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    expect(missingEmbedded.spans[0]?.font.source).toBe('generic');
    expect(missingEmbedded.diagnostics[0]?.message).toMatch(/unavailable/i);
    expect(missingGoogle.spans[0]?.font.source).toBe('generic');

    const carlito = { family: 'Carlito', weight: '400', style: 'normal', status: 'loaded' } as FontFace;
    const loaded = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [],
      googleFaces: [carlito],
      useGoogleFonts: true,
    });
    const substituted = loaded.text.shape({ text: 'x', fontSizePt: 10, fonts: { ascii: 'Calibri' } });
    expect(substituted.spans[0]?.font).toMatchObject({ source: 'substitute', resolvedFamily: 'Carlito' });
    expect(substituted.diagnostics[0]?.message).toMatch(/implementation-dependent/i);
  });

  it('requires loaded status and an exact family/weight/style match for every face', () => {
    const doc = model({
      embeddedFonts: [
        { fontName: 'Partial Embedded', partPath: 'word/fonts/regular.odttf', fontKey: '', style: 'regular' },
        { fontName: 'Partial Embedded', partPath: 'word/fonts/bold.odttf', fontKey: '', style: 'bold' },
      ],
    });
    const services = createLayoutServices(doc, {
      measureContext: measureContext(),
      embeddedFaces: [
        { family: '"Partial Embedded"', weight: '400', style: 'normal', status: 'loaded' },
        { family: 'Partial Embedded', weight: '700', style: 'normal', status: 'error' },
        { family: 'Timed Out', weight: '400', style: 'normal', status: 'loading' },
      ] as FontFace[],
    });
    const shape = (family: string, weight: number, style: 'normal' | 'italic' = 'normal') =>
      services.text.shape({ text: 'x', fontSizePt: 10, weight, style, fonts: { ascii: family } });

    expect(shape('Partial Embedded', 400).spans[0]?.font)
      .toMatchObject({ source: 'embedded', resolvedFamily: 'Partial Embedded' });
    expect(shape('Partial Embedded', 700).spans[0]?.font.source).toBe('generic');
    expect(shape('Partial Embedded', 400, 'italic').spans[0]?.font.source).toBe('generic');
    expect(shape('Timed Out', 400).spans[0]?.font.source).toBe('generic');
  });

  it('collects every currently representable math story, including nested tables', () => {
    const math = (value: string) => ({
      type: 'math', nodes: [{ type: 'text', text: value }], display: false, fontSize: 10,
    });
    const paragraph = (value: string) => ({ type: 'paragraph', runs: [math(value)] });
    const table = (value: string) => ({
      type: 'table', rows: [{ cells: [{ content: [paragraph(value)] }] }],
    });
    const doc = model({
      body: [paragraph('body')],
      headers: { default: { body: [table('header')] }, first: null, even: null },
      footers: { default: null, first: { body: [paragraph('footer')] }, even: null },
      footnotes: [{ id: '1', content: [table('footnote')] }],
      endnotes: [{ id: '2', content: [paragraph('endnote')] }],
    } as unknown as Partial<DocxDocumentModel>);
    const services = createLayoutServices(doc, { measureContext: measureContext() });

    for (const value of ['body', 'header', 'footer', 'footnote', 'endnote']) {
      const lookupKey = mathAstResourceKey({ nodes: [{ type: 'text', text: value }], display: false });
      expect(() => services.math.resolve(lookupKey), value).not.toThrow();
    }
  });

  it('requires runtime math handles to match available metadata exactly', () => {
    const available = {
      resourceKey: 'math:available', widthEm: 1, ascentEm: 0.8, descentEm: 0.2, diagnostics: [],
    };
    const unavailable = {
      resourceKey: 'math:unavailable', widthEm: 0, ascentEm: 0, descentEm: 0,
      available: false as const, diagnostics: [],
    };
    const drawable = {} as CanvasImageSource;

    expect(() => createLayoutServices(model(), { mathResources: [available], mathDrawables: new Map() }))
      .toThrow(/math.*membership|missing/i);
    expect(() => createLayoutServices(model(), {
      mathResources: [unavailable],
      mathDrawables: new Map([['math:unavailable', drawable]]),
    })).toThrow(/math.*membership|extra/i);

    const services = createLayoutServices(model(), {
      mathResources: [available, unavailable],
      mathDrawables: new Map([['math:available', drawable]]),
    });
    expect(privateResourceLookupOf(services)?.keys).toEqual(['math:available']);
  });

  it('gives main and worker factories identical fingerprints for identical successful snapshots', () => {
    const embedded = { family: 'Embedded', weight: '700', style: 'italic', status: 'loaded' } as FontFace;
    const options = {
      measureContext: measureContext(),
      embeddedFaces: [embedded],
      googleFaces: [] as FontFace[],
      localMetrics: { authored: { family: '__local_authored', lineHeightRatio: 1.25 } },
    };
    const main = createLayoutServices(model(), options);
    const worker = createLayoutServices(model(), { ...options, measureContext: measureContext() });

    expect(main.text.fingerprint).toBe(worker.text.fingerprint);
    expect(main.images.fingerprint).toBe(worker.images.fingerprint);
    expect(main.math.fingerprint).toBe(worker.math.fingerprint);
  });
});
