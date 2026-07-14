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
  { requestedFamily: 'Embedded Sans', resolvedFamily: 'Embedded Sans', source: 'embedded' },
  { requestedFamily: 'Meiryo', resolvedFamily: '__ooxml_local_meiryo', source: 'local' },
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
    expect(resolver.resolve({ requestedFamily: 'Calibri' }))
      .toMatchObject({ source: 'google', resolvedFamily: 'Carlito' });

    const substituted = resolver.resolve({ requestedFamily: 'Legacy Arabic' });
    expect(substituted).toMatchObject({ source: 'substitute', resolvedFamily: 'Noto Naskh Arabic' });
    expect(substituted.diagnostics[0]?.message).toMatch(/implementation-dependent font substitution/i);

    const missing = resolver.resolve({ requestedFamily: 'Missing Face', genericFamily: 'serif' });
    expect(missing).toMatchObject({ source: 'generic', resolvedFamily: 'serif' });
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
        ['国', 'local', '__ooxml_local_meiryo'],
        ['ش', 'substitute', 'Noto Naskh Arabic'],
      ]);
    expect(result.advancePt).toBe(30);
    expect(measured.map(({ resolvedFamily, weight, style }) => [resolvedFamily, weight, style]))
      .toEqual([
        ['Embedded Sans', 700, 'italic'],
        ['__ooxml_local_meiryo', 700, 'italic'],
        ['Noto Naskh Arabic', 700, 'italic'],
      ]);
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
});
