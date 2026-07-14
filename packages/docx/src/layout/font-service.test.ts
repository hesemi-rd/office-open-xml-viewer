import { describe, expect, it } from 'vitest';
import {
  createFontResolver,
  type FontInventoryFace,
} from './font-service.js';
import {
  createTextLayoutService,
  type GlyphMeasureRequest,
  type GlyphMeasurement,
} from './text.js';

const faces: readonly FontInventoryFace[] = [
  { requestedFamily: 'Embedded Sans', resolvedFamily: 'Embedded Sans', source: 'embedded', weights: [400, 700], styles: ['normal', 'italic'] },
  { requestedFamily: 'Meiryo', resolvedFamily: '__ooxml_local_meiryo', source: 'local', weights: [400], styles: ['normal'] },
  { requestedFamily: 'Calibri', resolvedFamily: 'Carlito', source: 'google' },
  { requestedFamily: 'Legacy Arabic', resolvedFamily: 'Noto Naskh Arabic', source: 'substitute' },
  { requestedFamily: 'Theme Serif', resolvedFamily: 'Theme Serif', source: 'local' },
];

describe('font layout services', () => {
  it('records embedded, local, Google, substitute, and generic resolution', () => {
    const resolver = createFontResolver(faces);

    expect(resolver.resolve({ requestedFamily: 'Embedded Sans', weight: 700, style: 'italic' }))
      .toMatchObject({ source: 'embedded', resolvedFamily: 'Embedded Sans', weight: 700, style: 'italic', diagnostics: [] });
    expect(resolver.resolve({ requestedFamily: 'Meiryo' }))
      .toMatchObject({ source: 'local', resolvedFamily: '__ooxml_local_meiryo' });
    expect(resolver.resolve({ requestedFamily: 'Meiryo', weight: 700 }))
      .toMatchObject({ source: 'generic', resolvedFamily: 'sans-serif', genericFamily: 'sans-serif' });
    expect(resolver.resolve({ requestedFamily: 'Calibri' }))
      .toMatchObject({ source: 'google', resolvedFamily: 'Carlito' });

    const substituted = resolver.resolve({ requestedFamily: 'Legacy Arabic' });
    expect(substituted).toMatchObject({ source: 'substitute', resolvedFamily: 'Noto Naskh Arabic' });
    expect(substituted.diagnostics[0]?.message).toMatch(/implementation-dependent font substitution/i);

    const missing = resolver.resolve({ requestedFamily: 'Missing Face', genericFamily: 'serif' });
    expect(missing).toMatchObject({ source: 'generic', resolvedFamily: 'serif', genericFamily: 'serif' });
    expect(missing.diagnostics[0]?.message).toContain('Missing Face');
  });

  it('selects font slots per script and shapes through the injected measurer', () => {
    const measured: GlyphMeasureRequest[] = [];
    const measure = (request: Readonly<GlyphMeasureRequest>): GlyphMeasurement => {
      measured.push({ ...request });
      return {
        advancePt: [...request.text].length * request.fontSizePt,
        ascentPt: request.fontSizePt * 0.8,
        descentPt: request.fontSizePt * 0.2,
      };
    };
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: { fingerprint: 'fake-glyphs-v1', measure },
    });

    const result = service.shape({
      text: 'A国ش',
      fontSizePt: 10,
      weight: 700,
      style: 'italic',
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Embedded Sans',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    });

    expect(result.spans.map((span) => [span.text, span.font.source, span.font.resolvedFamily]))
      .toEqual([
        ['A', 'embedded', 'Embedded Sans'],
        ['国', 'generic', 'sans-serif'],
        ['ش', 'embedded', 'Embedded Sans'],
      ]);
    expect(result.advancePt).toBe(30);
    expect(measured.map(({ resolvedFamily, weight, style }) => [resolvedFamily, weight, style]))
      .toEqual([
        ['Embedded Sans', 700, 'italic'],
        ['sans-serif', 700, 'italic'],
        ['Embedded Sans', 700, 'italic'],
      ]);

    const forced = service.shape({
      text: 'ش-12',
      fontSizePt: 10,
      complexScript: true,
      fonts: { ascii: 'Embedded Sans', complexScript: 'Legacy Arabic' },
    });
    expect(forced.spans).toHaveLength(1);
    expect(forced.spans[0]).toMatchObject({ text: 'ش-12', script: 'complexScript' });
    expect(forced.spans[0]?.font).toMatchObject({ source: 'substitute', resolvedFamily: 'Noto Naskh Arabic' });
  });

  it('uses theme slots without collapsing mixed-script runs to one family', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'fake-glyphs-v1',
        measure: (request) => ({
          advancePt: request.text.length,
          ascentPt: 1,
          descentPt: 0,
        }),
      },
    });

    const result = service.shape({
      text: 'A国',
      fontSizePt: 12,
      fonts: {},
      themeFonts: { ascii: 'Theme Serif', eastAsia: 'Meiryo' },
    });

    expect(result.spans.map((span) => span.font.requestedFamily))
      .toEqual(['Theme Serif', 'Meiryo']);
    expect(service.fingerprint).toMatch(/^text:/);
  });

  it('produces identical main and worker fingerprints from identical snapshots', () => {
    const make = () => createTextLayoutService({
      fonts: createFontResolver([...faces].reverse()),
      measurer: {
        fingerprint: 'canvas-text-metrics-v1',
        measure: () => ({ advancePt: 1, ascentPt: 1, descentPt: 0 }),
      },
    });

    const main = make();
    const worker = make();
    expect(main.fingerprint).toBe(worker.fingerprint);
    expect(main.shape({ text: 'A国', fontSizePt: 10, fonts: { ascii: 'Calibri', eastAsia: 'Meiryo' } }))
      .toEqual(worker.shape({ text: 'A国', fontSizePt: 10, fonts: { ascii: 'Calibri', eastAsia: 'Meiryo' } }));
  });

  it('keeps grapheme clusters intact across ECMA-376 script-slot boundaries', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'cluster-measurer-v1',
        measure: (request) => ({ advancePt: [...request.text].length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const shape = (text: string) => service.shape({
      text,
      fontSizePt: 10,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Theme Serif',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    });

    expect(shape('a\u0301').spans.map((span) => [span.text, span.script]))
      .toEqual([['a\u0301', 'ascii']]);
    expect(shape('国\u{E0100}').spans.map((span) => [span.text, span.script]))
      .toEqual([['国\u{E0100}', 'eastAsia']]);
    expect(shape('\u{20000}').spans.map((span) => [span.text, span.script]))
      .toEqual([['\u{20000}', 'eastAsia']]);
    expect(shape('👩‍💻').spans.map((span) => span.text)).toEqual(['👩‍💻']);
    expect(shape('ش-12').spans.every((span) => span.script !== 'complexScript')).toBe(true);
  });

  it('implements the complete ECMA-376 §17.3.2.26 character-range table', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'script-slot-table-v1',
        measure: (request) => ({ advancePt: [...request.text].length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const slot = (text: string, complexScript = false) => service.shape({
      text,
      fontSizePt: 10,
      complexScript,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Theme Serif',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    }).spans.map((span) => span.script);

    const cases: Array<[string, string, 'ascii' | 'highAnsi' | 'eastAsia']> = [
      ['ASCII', 'A', 'ascii'],
      ['Latin-1', 'é', 'highAnsi'],
      ['Hebrew', '\u05D0', 'ascii'],
      ['Arabic', '\u0627', 'ascii'],
      ['Syriac', '\u0710', 'ascii'],
      ['Thaana', '\u0780', 'ascii'],
      ['Hebrew presentation form', '\uFB1D', 'ascii'],
      ['Arabic presentation form A', '\uFB50', 'ascii'],
      ['Arabic presentation form B', '\uFE70', 'ascii'],
      ['Hangul Jamo', '\u1100', 'eastAsia'],
      ['Yi syllable', '\uA000', 'eastAsia'],
      ['CJK compatibility form', '\uFE30', 'eastAsia'],
      ['small form variant', '\uFE50', 'eastAsia'],
    ];
    for (const [name, text, expected] of cases) {
      expect(slot(text), name).toEqual([expected]);
    }
    expect(slot('\u05D0', true)).toEqual(['complexScript']);
    expect(slot('\u0627', true)).toEqual(['complexScript']);
  });

  it('implements the conditional eastAsia hint, language, charset, and cs precedence rows', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver(faces),
      measurer: {
        fingerprint: 'conditional-script-slot-table-v1',
        measure: (request) => ({ advancePt: [...request.text].length, ascentPt: 1, descentPt: 0 }),
      },
    });
    const slot = (
      text: string,
      options: {
        hint?: 'default' | 'eastAsia' | 'cs';
        eastAsiaLanguage?: string;
        eastAsiaFontCharset?: string;
        complexScript?: boolean;
      } = {},
    ) => service.shape({
      text,
      fontSizePt: 10,
      fontHint: options.hint,
      eastAsiaLanguage: options.eastAsiaLanguage,
      eastAsiaFontCharset: options.eastAsiaFontCharset,
      complexScript: options.complexScript,
      fonts: {
        ascii: 'Embedded Sans',
        highAnsi: 'Theme Serif',
        eastAsia: 'Meiryo',
        complexScript: 'Legacy Arabic',
      },
    }).spans[0]?.script;

    const eastAsia = { hint: 'eastAsia' as const };
    expect(slot('\u00a1', eastAsia)).toBe('eastAsia');
    expect(slot('\u00e0', { ...eastAsia, eastAsiaLanguage: 'zh-CN' })).toBe('eastAsia');
    expect(slot('\u00e0', { ...eastAsia, eastAsiaLanguage: 'ja-JP' })).toBe('highAnsi');
    expect(slot('\u0100', { ...eastAsia, eastAsiaLanguage: 'zh-TW' })).toBe('eastAsia');
    expect(slot('\u0180', { ...eastAsia, eastAsiaFontCharset: '88' })).toBe('eastAsia');
    expect(slot('\u0250', { ...eastAsia, eastAsiaFontCharset: '86' })).toBe('eastAsia');
    expect(slot('\u0100', { ...eastAsia, eastAsiaFontCharset: '80' })).toBe('highAnsi');
    for (const scalar of ['\u02b0', '\u0300', '\u0370', '\u0400', '\u2000', '\u20a0', '\u2100', '\u2190', '\u2200', '\u2300', '\u2400', '\u2500', '\u2600', '\u2700', '\ue000', '\ufb00']) {
      expect(slot(scalar, eastAsia), scalar).toBe('eastAsia');
      expect(slot(scalar), scalar).toBe('highAnsi');
    }
    expect(slot('\u1e00', { ...eastAsia, eastAsiaLanguage: 'zh-Hans' })).toBe('eastAsia');
    expect(slot('\u1e00', { ...eastAsia, eastAsiaLanguage: 'ko-KR' })).toBe('highAnsi');

    // Step 2: cs/rtl wins unless step 1 selected eastAsia while hint=eastAsia.
    expect(slot('\u4e00', { complexScript: true })).toBe('complexScript');
    expect(slot('\u4e00', { ...eastAsia, complexScript: true })).toBe('eastAsia');
    expect(slot('A', { ...eastAsia, complexScript: true })).toBe('complexScript');
  });

  it('includes immutable local geometry metrics in the text fingerprint', () => {
    const make = (lineHeightRatio: number) => createTextLayoutService({
      fonts: createFontResolver(faces),
      localMetrics: {
        meiryo: { family: '__ooxml_local_meiryo', lineHeightRatio },
      },
      measurer: {
        fingerprint: 'metrics-v1',
        measure: () => ({ advancePt: 1, ascentPt: 1, descentPt: 0 }),
      },
    });

    expect(make(1.3).fingerprint).not.toBe(make(1.31).fingerprint);
    expect(Object.isFrozen(make(1.3).localMetrics.meiryo)).toBe(true);
  });
});
