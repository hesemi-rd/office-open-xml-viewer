import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocTable,
  DocxDocumentModel,
  HeadersFooters,
  ShapeRun,
  ShapeText,
} from './types.js';

/** One rendered string and every authored font family that can supply it.
 * Empty text records are intentional: paragraph marks and drawing anchors can
 * affect line metrics even when they paint no glyphs. */
export interface DocxRenderedTextUsage {
  text: string;
  fontFamilies: readonly (string | null | undefined)[];
}

function* shapeTextUsages(shape: ShapeRun): Generator<DocxRenderedTextUsage> {
  if (shape.textPath) {
    yield { text: shape.textPath.string, fontFamilies: [shape.textPath.fontFamily] };
  }
  for (const block of shape.textBlocks ?? []) {
    yield* shapeBlockUsages(block);
  }
}

function* shapeBlockUsages(block: ShapeText): Generator<DocxRenderedTextUsage> {
  if (block.numbering) {
    yield {
      text: block.numbering.text,
      fontFamilies: [block.numbering.fontFamily, block.numbering.fontFamilyEastAsia],
    };
  }
  if (block.runs?.length) {
    for (const run of block.runs) {
      yield {
        text: run.text,
        // A run without an explicit axis inherits the block-level face.
        fontFamilies: [
          run.fontFamily,
          run.fontFamilyEastAsia,
          block.fontFamily,
        ],
      };
    }
  } else {
    yield { text: block.text, fontFamilies: [block.fontFamily] };
  }
}

function* runUsages(run: DocRun): Generator<DocxRenderedTextUsage> {
  if (run.type === 'text') {
    yield {
      text: run.text,
      fontFamilies: [run.fontFamily, run.fontFamilyEastAsia, run.fontFamilyCs],
    };
  } else if (run.type === 'field') {
    yield { text: run.fallbackText, fontFamilies: [run.fontFamily] };
  } else if (run.type === 'shape') {
    yield* shapeTextUsages(run);
  } else if (run.type === 'anchorHost') {
    yield {
      text: '',
      fontFamilies: [run.fontFamily, run.fontFamilyEastAsia],
    };
  }
}

function* paragraphUsages(paragraph: DocParagraph): Generator<DocxRenderedTextUsage> {
  // Empty paragraphs still reserve the resolved paragraph-mark line box.
  yield {
    text: '',
    fontFamilies: [paragraph.defaultFontFamily, paragraph.defaultFontFamilyEastAsia],
  };
  if (paragraph.numbering) {
    yield {
      text: paragraph.numbering.text,
      fontFamilies: [
        paragraph.numbering.fontFamily,
        paragraph.numbering.fontFamilyEastAsia,
      ],
    };
  }
  for (const run of paragraph.runs) yield* runUsages(run);
}

function* tableUsages(table: DocTable): Generator<DocxRenderedTextUsage> {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      yield* bodyUsages(cell.content as BodyElement[]);
    }
  }
}

function* headerFooterUsages(
  stories: HeadersFooters | null | undefined,
): Generator<DocxRenderedTextUsage> {
  if (!stories) return;
  for (const story of [stories.default, stories.first, stories.even]) {
    if (story) yield* bodyUsages(story.body);
  }
}

function* bodyUsages(body: readonly BodyElement[]): Generator<DocxRenderedTextUsage> {
  for (const element of body) {
    if (element.type === 'paragraph') {
      yield* paragraphUsages(element);
    } else if (element.type === 'table') {
      yield* tableUsages(element);
    } else if (element.type === 'sectionBreak') {
      // Non-final sections keep their resolved header/footer stories on the
      // marker; the top-level sets represent only the final section.
      yield* headerFooterUsages(element.headers);
      yield* headerFooterUsages(element.footers);
    }
  }
}

/** Traverse every rendered DOCX story once. This is shared by script-aware web
 * font preloading and exact-local metric discovery so those paths cannot drift
 * on nested tables, section headers/footers, notes, or drawing text. Comments
 * are excluded because the page renderer does not paint comment bodies. */
export function* docxRenderedTextUsages(
  doc: DocxDocumentModel,
): Generator<DocxRenderedTextUsage> {
  yield* bodyUsages(doc.body ?? []);
  yield* headerFooterUsages(doc.headers);
  yield* headerFooterUsages(doc.footers);
  for (const note of [...(doc.footnotes ?? []), ...(doc.endnotes ?? [])]) {
    yield* bodyUsages(note.content);
  }
}

/** Unique authored families in first-rendered-use order. */
export function docxRenderedFontFamilies(doc: DocxDocumentModel): string[] {
  const families = new Set<string>();
  for (const usage of docxRenderedTextUsages(doc)) {
    for (const family of usage.fontFamilies) {
      const trimmed = family?.trim();
      if (trimmed) families.add(trimmed);
    }
  }
  return [...families];
}
