import { describe, it, expect } from 'vitest';
import { layoutRichTextLines } from './renderer.js';
import type { CellFont, Run } from './types.js';

// ECMA-376 §18.8.1 (CT_CellAlignment) @wrapText: text in a cell is line-wrapped
// within the cell. Cell rich text is stored as <r>/<t> runs (§18.4.4 / §18.4.12)
// with xml:space="preserve", so a hard line break authored with Alt+Enter is a
// literal LF (U+000A) inside the run text. Each rendered line — INCLUDING a blank
// line produced by consecutive breaks, or a leading/trailing break — reserves one
// single-line height, exactly as Excel renders it and as the shape-text path now
// does (PR #583, the xlsx sibling of docx PR #582).
//
// `layoutRichTextLines` flushes a line region at every LF, but its `flush` helper
// early-returned when the region was empty (`if (cur.length === 0) return;`), so a
// blank line was silently dropped and every line below it pulled up by one line
// height. This is the cell rich-text analog of the same blank-line-collapse bug.
//
// The cell path derives line height analytically (font size × 1.2 via vMetricPx),
// not from font metrics, so the bug is visible without an asymmetric font stub: a
// blank line must contribute the SAME 1.2-em height as a text line of the same
// effective font size — here exposed as `RichLine.maxFontSize`, the per-line
// height source the renderer feeds to `vMetricPx`.

const BASE: CellFont = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: null,
};

function makeCtx(): CanvasRenderingContext2D {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  return {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

/** Lay out a single-run cell value. A very wide maxWidth disables soft-wrapping
 *  so only explicit LF breaks split lines — isolating the blank-line behaviour. */
function layout(text: string, font?: Run['font']) {
  return layoutRichTextLines(makeCtx(), [{ text, font }] as Run[], BASE, 1, 100000);
}

describe('empty line height in cell rich text (§18.8.1 wrapText / §18.4.4 r)', () => {
  it('a blank line between two text lines reserves one single-line height', () => {
    // Reference: "A\nB" — two single-line regions, no blank between.
    const ctrl = layout('A\nB');
    expect(ctrl.length).toBe(2);
    expect(ctrl[0].segments.length).toBe(1);
    expect(ctrl[1].segments.length).toBe(1);

    // Subject: "A\n\nB" — the consecutive LFs create a blank middle line. It must
    // be emitted (3 lines, not 2) so the "B" line sits one line height lower.
    const subj = layout('A\n\nB');
    expect(subj.length).toBe(3);
    // The middle line is blank (no drawn segments) ...
    expect(subj[1].segments.length).toBe(0);
    // ... yet reserves the SAME single-line height as a text line, not 0.
    expect(subj[1].maxFontSize).toBe(11);
    expect(subj[1].maxFontSize).toBe(subj[0].maxFontSize);
  });

  it('the blank line is sized from the effective run font size, not a constant or 0', () => {
    const r = layout('A\n\nB', {
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      size: 22,
    });
    expect(r.length).toBe(3);
    expect(r[1].segments.length).toBe(0);
    // 22pt run → the blank line reserves a 22pt line height, not the 11pt default.
    expect(r[1].maxFontSize).toBe(22);
  });

  it('a leading blank line (value begins with a break) is preserved', () => {
    const r = layout('\nA');
    expect(r.length).toBe(2);
    expect(r[0].segments.length).toBe(0);
    expect(r[0].maxFontSize).toBe(11);
    expect(r[1].segments.length).toBe(1);
  });

  it('a trailing blank line (value ends with a break) is preserved', () => {
    const r = layout('A\n');
    expect(r.length).toBe(2);
    expect(r[0].segments.length).toBe(1);
    expect(r[1].segments.length).toBe(0);
    expect(r[1].maxFontSize).toBe(11);
  });

  it('an empty rich-text value produces no fabricated line', () => {
    // A value with no characters and no breaks contributes nothing — the renderer
    // must not invent a blank line out of an empty cell.
    expect(layout('').length).toBe(0);
  });
});
