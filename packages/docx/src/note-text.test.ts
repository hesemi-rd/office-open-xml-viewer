import { describe, it, expect } from 'vitest';
import { noteText } from './types.js';
import type { DocNote, DocParagraph, DocxTextRun } from './types.js';

// noteText flattens a footnote/endnote's content to plain text, dropping the
// auto-number reference marker (a `noteRef` run) and joining paragraphs with a
// space. This is the data-only projection of ECMA-376 §17.11 note content.

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: null, isLink: false,
    background: null, vertAlign: null, hyperlink: null,
    ...extra,
  };
}

function para(runs: DocxTextRun[]): DocParagraph {
  return {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r })) as DocParagraph['runs'],
  };
}

describe('noteText', () => {
  it('drops the leading footnoteRef marker and returns the body text', () => {
    const note: DocNote = {
      id: '1',
      content: [
        { type: 'paragraph', ...para([
          textRun('1', { noteRef: { kind: 'footnote', id: '' }, vertAlign: 'super' }),
          textRun(' On forest age and longevity.'),
        ]) },
      ],
    };
    expect(noteText(note)).toBe('On forest age and longevity.');
  });

  it('joins multiple paragraphs with a single space', () => {
    const note: DocNote = {
      id: '2',
      content: [
        { type: 'paragraph', ...para([textRun('First line.')]) },
        { type: 'paragraph', ...para([textRun('Second line.')]) },
      ],
    };
    expect(noteText(note)).toBe('First line. Second line.');
  });

  it('returns an empty string for an empty note', () => {
    expect(noteText({ id: '3', content: [] })).toBe('');
  });
});
