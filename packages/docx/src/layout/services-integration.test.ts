import { describe, expect, it } from 'vitest';
import { buildSegments, layoutLines, type LineLayoutEnvironment } from '../line-layout.js';
import { createLayoutServices } from '../renderer.js';
import type { DocRun, DocxDocumentModel } from '../types.js';
import type { TextLayoutService } from './text.js';

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
    layoutLines(measureContext(), segments, 300, 0, 1);

    expect(afterSegmentation).toBeGreaterThanOrEqual(2);
    expect(calls).toBeGreaterThan(afterSegmentation);
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
