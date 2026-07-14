import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxDocumentModel,
  DocxTextRun,
  FieldRun,
  HeadersFooters,
  NumberingInfo,
} from './types.js';
import type { NumberingMarkerShapeInput, SourceRef } from './layout/types.js';
import type { MathOccurrence } from './layout/resources.js';
import { mathResourceKey } from './layout/source-key.js';
import type { TextFontSlotPresence, TextFontSlots } from './layout/text.js';

export interface InternalRunFontSlots {
  readonly direct: TextFontSlots;
  readonly theme: TextFontSlots;
  readonly themePresent: TextFontSlotPresence;
}

/** Parser-emitted metadata intentionally kept outside the stable public model.
 * Ordinary text and field results share these resolved WordprocessingML axes. */
export interface InternalRunSlotMetadata {
  fontFamilyHighAnsi?: string | null;
  fontSlots?: InternalRunFontSlots;
  fontFamilyEastAsia?: string | null;
  fontHint?: 'default' | 'eastAsia' | 'cs';
  rtl?: boolean;
  cs?: boolean;
  fontFamilyCs?: string | null;
  fontSizeCs?: number;
  boldCs?: boolean;
  italicCs?: boolean;
  langBidi?: string;
  langEastAsia?: string;
}

/** Effective parser-owned run facts used by non-content glyphs such as list
 * markers and paragraph marks. Kept off the stable public document model. */
export interface InternalRunFontFacts extends InternalRunSlotMetadata {
  fontFamily?: string | null;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  kerning?: number;
}

export interface InternalNumberingInfo extends NumberingInfo {
  fontFacts?: InternalRunFontFacts;
}

export interface InternalDocParagraph extends DocParagraph {
  numbering: InternalNumberingInfo | null;
  paragraphMarkFontFacts?: InternalRunFontFacts;
}

type TextOnlyMetadata = Pick<
  DocxTextRun,
  | 'ruby' | 'revision' | 'hyperlink' | 'hyperlinkAnchor'
  | 'underlineStyle' | 'underlineColor' | 'colorAuto' | 'border'
  | 'snapToGrid' | 'charSpacing' | 'charScale' | 'fitTextVal' | 'fitTextId'
  | 'position' | 'kerning' | 'eastAsianVert' | 'eastAsianVertCompress'
>;

export type InternalTextRun = DocxTextRun & InternalRunSlotMetadata;
export type InternalFieldRun = FieldRun & Partial<TextOnlyMetadata> & InternalRunSlotMetadata;
export type InternalTextBearingRun = InternalTextRun | InternalFieldRun;
export type InternalMathRun = Extract<DocRun, { type: 'math' }> & {
  readonly source: SourceRef;
  readonly resourceKey: string;
};

export interface InternalDocxDocumentModel extends DocxDocumentModel {
  fontFamilyCharsets?: Record<string, string>;
}

export interface NormalizedDocumentInput {
  readonly document: InternalDocxDocumentModel;
  readonly mathOccurrences: readonly MathOccurrence[];
}

/** Snapshot the parser's effective numbering-level rPr into the plain retained
 * layout contract. This is the parser-model/layout boundary: layout code never
 * dereferences the private parser extension itself. */
export function numberingMarkerShapeInput(
  num: NumberingInfo,
  fallbackFontSizePt: number,
): NumberingMarkerShapeInput {
  const facts = internalNumberingInfo(num).fontFacts;
  const complexScript = facts?.rtl === true || facts?.cs === true;
  const fontSizePt = complexScript
    ? (facts?.fontSizeCs ?? facts?.fontSize ?? fallbackFontSizePt)
    : (facts?.fontSize ?? fallbackFontSizePt);
  const ascii = facts?.fontFamily ?? num.fontFamily ?? null;
  const fallbackFonts: TextFontSlots = {
    ascii,
    highAnsi: facts?.fontFamilyHighAnsi ?? ascii,
    eastAsia: facts?.fontFamilyEastAsia ?? num.fontFamilyEastAsia ?? ascii,
    complexScript: facts?.fontFamilyCs ?? ascii,
  };
  const slots = facts?.fontSlots;
  return Object.freeze({
    fontSizePt,
    fonts: Object.freeze({ ...(slots?.direct ?? fallbackFonts) }),
    themeFonts: slots?.theme ? Object.freeze({ ...slots.theme }) : undefined,
    themeFontPresence: slots?.themePresent
      ? Object.freeze({ ...slots.themePresent })
      : undefined,
    weight: (complexScript ? (facts?.boldCs ?? false) : (facts?.bold ?? false)) ? 700 : 400,
    style: (complexScript ? (facts?.italicCs ?? false) : (facts?.italic ?? false))
      ? 'italic'
      : 'normal',
    complexScript,
    fontHint: facts?.fontHint,
    eastAsiaLanguage: facts?.langEastAsia,
    kerning: facts?.kerning == null ? undefined : fontSizePt >= facts.kerning,
  });
}

/** Pure structural normalization for stable math addressing. Only ancestry that
 * contains a math run is shallow-cloned; the caller's parser model is untouched. */
export function normalizeInternalDocumentModel(doc: DocxDocumentModel): NormalizedDocumentInput {
  const occurrences: MathOccurrence[] = [];
  const normalizeBody = (
    body: BodyElement[],
    story: SourceRef['story'],
    storyInstance: string,
    prefix: number[] = [],
  ): BodyElement[] => {
    let changed = false;
    const normalized = body.map((element, elementIndex): BodyElement => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        let runsChanged = false;
        const runs = element.runs.map((run, runIndex): DocRun => {
          if (run.type !== 'math') return run;
          runsChanged = true;
          const source: SourceRef = Object.freeze({
            story,
            storyInstance,
            path: Object.freeze([...path, runIndex]),
          });
          const resourceKey = mathResourceKey(source, run.display ? 'display' : 'inline');
          occurrences.push(Object.freeze({
            nodes: run.nodes,
            display: run.display,
            source,
            resourceKey,
          }));
          return Object.freeze({ ...run, source, resourceKey }) as InternalMathRun;
        });
        if (!runsChanged) return element;
        changed = true;
        return { ...element, runs };
      }
      if (element.type === 'table') {
        let tableChanged = false;
        const rows = element.rows.map((row, rowIndex) => {
          let rowChanged = false;
          const cells = row.cells.map((cell, cellIndex) => {
            const content = normalizeBody(
              cell.content as BodyElement[], story, storyInstance, [...path, rowIndex, cellIndex],
            );
            if (content === cell.content) return cell;
            rowChanged = true;
            return { ...cell, content: content as typeof cell.content };
          });
          if (!rowChanged) return row;
          tableChanged = true;
          return { ...row, cells };
        });
        if (!tableChanged) return element;
        changed = true;
        return { ...element, rows } as BodyElement;
      }
      if (element.type !== 'sectionBreak') return element;
      let sectionChanged = false;
      const normalizeParts = (
        parts: HeadersFooters | undefined,
        partStory: 'header' | 'footer',
      ): HeadersFooters | undefined => {
        if (!parts) return parts;
        let result = parts;
        for (const kind of ['default', 'first', 'even'] as const) {
          const part = parts[kind];
          if (!part) continue;
          const nextBody = normalizeBody(part.body, partStory, `section:${elementIndex}:${kind}`);
          if (nextBody === part.body) continue;
          if (result === parts) result = { ...parts };
          result[kind] = { ...part, body: nextBody };
          sectionChanged = true;
        }
        return result;
      };
      const headers = normalizeParts(element.headers, 'header');
      const footers = normalizeParts(element.footers, 'footer');
      if (!sectionChanged) return element;
      changed = true;
      return { ...element, headers, footers };
    });
    return changed ? normalized : body;
  };
  const normalizeParts = (
    parts: HeadersFooters,
    story: 'header' | 'footer',
  ): HeadersFooters => {
    let result = parts;
    for (const kind of ['default', 'first', 'even'] as const) {
      const part = parts[kind];
      if (!part) continue;
      const body = normalizeBody(part.body, story, kind);
      if (body === part.body) continue;
      if (result === parts) result = { ...parts };
      result[kind] = { ...part, body };
    }
    return result;
  };
  const body = normalizeBody(doc.body, 'body', 'body');
  const headers = normalizeParts(doc.headers, 'header');
  const footers = normalizeParts(doc.footers, 'footer');
  const normalizeNotes = <T extends { id: string; content: BodyElement[] }>(
    notes: T[] | undefined,
    story: 'footnote' | 'endnote',
  ): T[] | undefined => {
    if (!notes) return notes;
    let changed = false;
    const normalized = notes.map((note) => {
      const content = normalizeBody(note.content, story, note.id);
      if (content === note.content) return note;
      changed = true;
      return { ...note, content };
    });
    return changed ? normalized : notes;
  };
  const footnotes = normalizeNotes(doc.footnotes, 'footnote');
  const endnotes = normalizeNotes(doc.endnotes, 'endnote');
  const changed = body !== doc.body || headers !== doc.headers || footers !== doc.footers
    || footnotes !== doc.footnotes || endnotes !== doc.endnotes;
  return Object.freeze({
    document: (changed ? { ...doc, body, headers, footers, footnotes, endnotes } : doc) as InternalDocxDocumentModel,
    mathOccurrences: Object.freeze(occurrences),
  });
}

export function internalFieldRun(run: FieldRun): InternalFieldRun {
  return run as InternalFieldRun;
}

export function internalTextRun(run: DocxTextRun): InternalTextRun {
  return run as InternalTextRun;
}

export function internalNumberingInfo(numbering: NumberingInfo): InternalNumberingInfo {
  return numbering as InternalNumberingInfo;
}

export function internalParagraph(paragraph: DocParagraph): InternalDocParagraph {
  return paragraph as InternalDocParagraph;
}

export function internalDocumentModel(doc: DocxDocumentModel): InternalDocxDocumentModel {
  return doc as InternalDocxDocumentModel;
}
