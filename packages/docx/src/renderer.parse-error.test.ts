import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer';
import type { DocxDocumentModel, SectionProps } from './types';

/**
 * RB7: a document carrying `parseError` (a degraded `word/document.xml`, or —
 * since the MAJOR fix — a corrupt ZIP container) renders a visible placeholder
 * page instead of a blank white sheet. This drives {@link renderDocumentToCanvas}
 * against a recording 2D context and asserts the placeholder is painted (the
 * heading + the part-tagged message reach `fillText`), and that a healthy
 * document never takes that branch. pptx has the twin of this test
 * (`renderer.parse-error.test.ts`); docx lacked one.
 */

interface DrawCall {
  op: string;
  args: unknown[];
}

/** A minimal recording 2D context that logs the draw ops we assert on. */
function recordingCtx(): { ctx: CanvasRenderingContext2D; calls: DrawCall[] } {
  const calls: DrawCall[] = [];
  const rec =
    (op: string) =>
    (...args: unknown[]) => {
      calls.push({ op, args });
    };
  const ctx = {
    save: rec('save'),
    restore: rec('restore'),
    scale: rec('scale'),
    setTransform: rec('setTransform'),
    translate: rec('translate'),
    rotate: rec('rotate'),
    clip: rec('clip'),
    rect: rec('rect'),
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    fillRect: rec('fillRect'),
    strokeRect: rec('strokeRect'),
    fillText: rec('fillText'),
    setLineDash: rec('setLineDash'),
    beginPath: rec('beginPath'),
    measureText: (t: string) => ({ width: t.length * 6 }),
  } as unknown as CanvasRenderingContext2D;
  return { ctx, calls };
}

/** A canvas stub whose getContext returns the recording ctx. */
function stubCanvas(ctx: CanvasRenderingContext2D): HTMLCanvasElement {
  return {
    width: 0,
    height: 0,
    style: {} as CSSStyleDeclaration,
    offsetWidth: 816,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
}

/** US-letter section (pt), enough geometry for the renderer to size the page. */
const LETTER_SECTION: SectionProps = {
  pageWidth: 612,
  pageHeight: 792,
  marginTop: 72,
  marginRight: 72,
  marginBottom: 72,
  marginLeft: 72,
  headerDistance: 36,
  footerDistance: 36,
  titlePage: false,
  evenAndOddHeaders: false,
};

function emptyHeadersFooters() {
  return { default: null, first: null, even: null };
}

function degradedDoc(parseError: string): DocxDocumentModel {
  // Mirrors the Rust `degraded_document`: empty body, theme fonts, parseError set.
  return {
    section: LETTER_SECTION,
    body: [],
    headers: emptyHeadersFooters(),
    footers: emptyHeadersFooters(),
    parseError,
  };
}

describe('RB7 renderDocumentToCanvas placeholder', () => {
  it('paints a placeholder carrying the parseError message for a degraded document', async () => {
    const { ctx, calls } = recordingCtx();
    const canvas = stubCanvas(ctx);
    await renderDocumentToCanvas(
      degradedDoc('word/document.xml: unexpected end of stream'),
      canvas,
      0,
      { width: 816, dpr: 1 },
    );

    const texts = calls.filter((c) => c.op === 'fillText').map((c) => String(c.args[0]));
    // The heading and the part-tagged detail both reach the canvas.
    expect(texts.some((t) => t.includes('could not be displayed'))).toBe(true);
    expect(texts.join(' ')).toContain('document.xml');
    // A filled page + framed card were drawn.
    expect(calls.some((c) => c.op === 'fillRect')).toBe(true);
    expect(calls.some((c) => c.op === 'strokeRect')).toBe(true);
  });

  it('surfaces a corrupt-CONTAINER parseError (RB7 MAJOR) in the placeholder', async () => {
    const { ctx, calls } = recordingCtx();
    const canvas = stubCanvas(ctx);
    await renderDocumentToCanvas(
      degradedDoc('(zip container): invalid Zip archive: Could not find EOCD'),
      canvas,
      0,
      { width: 816, dpr: 1 },
    );
    const texts = calls.filter((c) => c.op === 'fillText').map((c) => String(c.args[0]));
    expect(texts.join(' ')).toContain('zip container');
  });

  it('a healthy document (no parseError) does NOT draw the placeholder heading', async () => {
    const { ctx, calls } = recordingCtx();
    const canvas = stubCanvas(ctx);
    const healthy: DocxDocumentModel = {
      section: LETTER_SECTION,
      body: [],
      headers: emptyHeadersFooters(),
      footers: emptyHeadersFooters(),
    };
    await renderDocumentToCanvas(healthy, canvas, 0, { width: 816, dpr: 1 });
    const texts = calls.filter((c) => c.op === 'fillText').map((c) => String(c.args[0]));
    expect(texts.some((t) => t.includes('could not be displayed'))).toBe(false);
  });
});
