import { describe, expect, it } from 'vitest';
import { docxFontPreloadNames } from './google-fonts.js';
import type { DocxDocumentModel } from './types.js';

/** Build a minimal model whose body is a single paragraph with one text run. */
function docWith(text: string, major = 'Calibri', minor = 'Calibri'): DocxDocumentModel {
  return {
    section: {} as DocxDocumentModel['section'],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    majorFont: major,
    minorFont: minor,
    body: [
      {
        type: 'paragraph',
        runs: [{ type: 'text', text } as never],
      } as never,
    ],
  } as DocxDocumentModel;
}

describe('docxFontPreloadNames — script-aware preload', () => {
  it('pure-Latin doc preloads ONLY the theme fonts (no CJK / script faces)', () => {
    const names = docxFontPreloadNames(docWith('Hello, world.'));
    expect(names).toEqual(['Calibri', 'Calibri']);
    // The expensive CJK faces must NOT be queued for a Latin document.
    expect(names).not.toContain('Noto Sans JP');
    expect(names).not.toContain('Noto Sans KR');
    expect(names).not.toContain('Noto Naskh Arabic');
  });

  it('Japanese doc preloads the JP Noto faces', () => {
    const names = docxFontPreloadNames(docWith('こんにちは世界'));
    expect(names).toContain('Noto Sans JP');
    expect(names).toContain('Noto Serif JP');
    expect(names).not.toContain('Noto Sans KR');
  });

  it('Han with a Korean theme font uses the kr lang hint', () => {
    const names = docxFontPreloadNames(docWith('漢字', 'Malgun Gothic', 'Malgun Gothic'));
    expect(names).toContain('Noto Sans KR');
    expect(names).not.toContain('Noto Sans JP');
  });

  it('is deterministic — same model yields the same set (main == worker)', () => {
    const doc = docWith('日本語 العربية');
    expect(docxFontPreloadNames(doc)).toEqual(docxFontPreloadNames(doc));
  });
});
