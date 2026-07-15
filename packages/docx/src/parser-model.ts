import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxDocumentModel,
  DocxTextRun,
  FieldRun,
  HeadersFooters,
  ImageRun,
  NumberingInfo,
  LineNumbering,
  ChartRun,
  ShapeRun,
  TextPath,
} from './types.js';
import type {
  NumberingMarkerShapeInput,
  SourceRef,
  VmlTextPathAcquisitionInput,
} from './layout/types.js';
import type { MathOccurrence } from './layout/resources.js';
import { anchorOccurrenceKey, mathResourceKey } from './layout/source-key.js';
import type { TextFontSlotPresence, TextFontSlots } from './layout/text.js';
import type { ParagraphAcquisitionInput, ParagraphAcquisitionRun } from './layout/text.js';
import type { AnchorAcquisitionInput, InternalAnchorRunWire } from './layout/anchor-input.js';
import {
  paragraphTypographyAcquisitionInput,
  runTypographyAcquisitionInput,
  type InternalParagraphTypographyWire,
  type InternalRunTypographyWire,
} from './layout/typography-input.js';
import { deepFreezePlainData, snapshotPlainData } from './layout/plain-data.js';
import {
  normalizeTextBoxInput,
  type NormalizedTextBoxParagraphInput,
} from './layout/textbox-input.js';

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

interface InternalSectionPlacementWire {
  readonly sectionId: string;
  readonly vAlign?: string | null;
  readonly lineNumbering?: LineNumbering | null;
}

type InternalSectionBreak = Extract<BodyElement, { type: 'sectionBreak' }> & {
  readonly __sectionPlacement?: InternalSectionPlacementWire;
};

/** Immutable parser-private section placement input. It is intentionally kept
 * outside BodyElement/DocxDocumentModel's stable declaration surface. */
export interface SectionPlacementInput {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
}

interface DocumentSectionPlacementInputs {
  readonly endingSections: ReadonlyMap<number, SectionPlacementInput>;
  readonly finalSection: SectionPlacementInput;
}

const sectionPlacementInputsByDocument = new WeakMap<object, DocumentSectionPlacementInputs>();
// The paginator seam receives two independent identity inputs. Nest the cache so
// reusing one body with a different final SectionProps cannot inherit stale final
// placement facts. Each entry remains an acquisition-time snapshot: subsequent
// caller mutation of the same SectionProps object does not rewrite retained data.
const sectionPlacementInputsByBody = new WeakMap<
  object,
  WeakMap<object, DocumentSectionPlacementInputs>
>();

function setBodySectionPlacementInputs(
  body: readonly BodyElement[],
  finalSection: DocxDocumentModel['section'] | undefined,
  inputs: DocumentSectionPlacementInputs,
): void {
  if (!finalSection || typeof finalSection !== 'object') return;
  let byFinalSection = sectionPlacementInputsByBody.get(body);
  if (!byFinalSection) {
    byFinalSection = new WeakMap<object, DocumentSectionPlacementInputs>();
    sectionPlacementInputsByBody.set(body, byFinalSection);
  }
  byFinalSection.set(finalSection, inputs);
}

function projectSectionPlacementInputs(doc: InternalDocxDocumentModel): DocumentSectionPlacementInputs {
  const endingSections = new Map<number, SectionPlacementInput>();
  let ordinal = 0;
  doc.body.forEach((element, bodyIndex) => {
    if (element.type !== 'sectionBreak') return;
    const wire = (element as InternalSectionBreak).__sectionPlacement;
    endingSections.set(bodyIndex, snapshotPlainData({
      sectionId: wire?.sectionId ?? `section:${ordinal}`,
      vAlign: wire?.vAlign ?? null,
      lineNumbering: wire?.lineNumbering ?? null,
    }, 'DOCX ending-section placement input'));
    ordinal += 1;
  });
  return Object.freeze({
    endingSections,
    finalSection: snapshotPlainData({
      sectionId: `section:${ordinal}`,
      // Resource-only entry points (for example image preloading) historically
      // accept a partial document projection with no section. Section placement
      // is irrelevant there, so preserve that compatibility with neutral facts.
      vAlign: doc.section?.vAlign ?? null,
      lineNumbering: doc.section?.lineNumbering ?? null,
    }, 'DOCX final-section placement input'),
  });
}

/** Resolve the section which owns body content beginning at `startIndex`.
 * Non-final section facts come from the next terminating SectionBreak; the
 * body-level sectPr owns the final section. */
export function sectionPlacementInputFrom(
  doc: InternalDocxDocumentModel,
  startIndex: number,
): SectionPlacementInput {
  let inputs = sectionPlacementInputsByDocument.get(doc);
  if (!inputs) {
    inputs = projectSectionPlacementInputs(doc);
    sectionPlacementInputsByDocument.set(doc, inputs);
  }
  for (let index = startIndex; index < doc.body.length; index += 1) {
    if (doc.body[index]?.type !== 'sectionBreak') continue;
    return inputs.endingSections.get(index) ?? inputs.finalSection;
  }
  return inputs.finalSection;
}

/** Body-array keyed twin used by the paginator, whose stable public signature
 * receives body + final SectionProps rather than the document wrapper. */
export function sectionPlacementInputFromBody(
  body: readonly BodyElement[],
  finalSection: DocxDocumentModel['section'],
  startIndex: number,
): SectionPlacementInput {
  let inputs = sectionPlacementInputsByBody.get(body)?.get(finalSection);
  if (!inputs) {
    const synthetic = { body, section: finalSection } as InternalDocxDocumentModel;
    inputs = projectSectionPlacementInputs(synthetic);
    setBodySectionPlacementInputs(body, finalSection, inputs);
  }
  for (let index = startIndex; index < body.length; index += 1) {
    if (body[index]?.type !== 'sectionBreak') continue;
    return inputs.endingSections.get(index) ?? inputs.finalSection;
  }
  return inputs.finalSection;
}

/** Resolved transitional VML facts emitted by the parser in addition to the
 * stable public `TextPath` surface. CT_Path owns `textPathOk`; CT_TextPath owns
 * the remaining switches. They stay private because they are acquisition
 * policy, not consumer-facing document content. */
export interface InternalVmlTextPath extends TextPath {
  textPathOk?: boolean;
  on?: boolean;
  fitShape?: boolean;
  fitPath?: boolean;
  trim?: boolean;
  xScale?: boolean;
  fontSizePt?: number;
}

export interface InternalShapeRun extends ShapeRun {
  textPath?: InternalVmlTextPath | null;
}

export interface NormalizedDocumentInput {
  readonly document: InternalDocxDocumentModel;
  readonly mathOccurrences: readonly MathOccurrence[];
}

/** Snapshot VML WordArt semantics at the parser/model boundary. The retained
 * acquisition layer consumes this clone-safe value and never needs to inspect
 * parser-only extensions on the public ShapeRun object. */
export function vmlTextPathAcquisitionInput(
  shape: Readonly<ShapeRun>,
): Readonly<VmlTextPathAcquisitionInput> | undefined {
  const textPath = (shape as Readonly<InternalShapeRun>).textPath;
  if (!textPath) return undefined;
  return snapshotPlainData({
    string: textPath.string,
    ...(textPath.fontFamily !== undefined ? { fontFamily: textPath.fontFamily } : {}),
    bold: textPath.bold ?? false,
    italic: textPath.italic ?? false,
    ...(textPath.textPathOk !== undefined ? { textPathOk: textPath.textPathOk } : {}),
    ...(textPath.on !== undefined ? { on: textPath.on } : {}),
    ...(textPath.fitShape !== undefined ? { fitShape: textPath.fitShape } : {}),
    ...(textPath.fitPath !== undefined ? { fitPath: textPath.fitPath } : {}),
    ...(textPath.trim !== undefined ? { trim: textPath.trim } : {}),
    ...(textPath.xScale !== undefined ? { xScale: textPath.xScale } : {}),
    ...(textPath.fontSizePt !== undefined ? { fontSizePt: textPath.fontSizePt } : {}),
  }, 'DOCX VML text path acquisition input');
}

/** Project parser-only anchor facts without widening the public run contract.
 * Parser-produced malformed input remains distinguishable from a hand-built
 * public run because required-but-missing values are explicit nulls. */
export function anchorAcquisitionInput(
  run: Readonly<DocRun | ShapeRun | ImageRun | ChartRun>,
): Readonly<AnchorAcquisitionInput> | undefined {
  const wire = (run as Readonly<DocRun & InternalAnchorRunWire>).__anchorAcquisition;
  if (wire === undefined) return undefined;
  return snapshotPlainData(wire, 'DOCX anchor acquisition input');
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

/** Project effective numbering-level run properties before a shape crosses the
 * parser/layout boundary. Public hand-built ShapeRun values use the normalizer's
 * compatibility fallback; parser-created shapes retain the full resolved slot
 * and theme facts without exposing their private wire object to layout. */
export function textBoxAcquisitionInput(
  shape: Readonly<ShapeRun>,
  source: SourceRef,
): readonly NormalizedTextBoxParagraphInput[] {
  return normalizeTextBoxInput(shape, source, numberingMarkerShapeInput);
}

/** Snapshot private paragraph-mark rPr facts at the parser boundary. Retained
 * line layout receives only this plain immutable service input. */
export function paragraphMarkShapeInput(
  paragraph: DocParagraph,
): NumberingMarkerShapeInput | undefined {
  const facts = internalParagraph(paragraph).paragraphMarkFontFacts;
  if (!facts) return undefined;
  const complexScript = facts.rtl === true || facts.cs === true;
  const fallbackFontSizePt = paragraph.runs.find(
    (run): run is Extract<DocRun, { type: 'text' | 'field' }> => run.type === 'text' || run.type === 'field',
  )?.fontSize ?? paragraph.defaultFontSize ?? 10;
  const fontSizePt = complexScript
    ? (facts.fontSizeCs ?? facts.fontSize ?? fallbackFontSizePt)
    : (facts.fontSize ?? fallbackFontSizePt);
  const ascii = facts.fontFamily ?? paragraph.defaultFontFamily ?? null;
  const fallbackFonts: TextFontSlots = {
    ascii,
    highAnsi: facts.fontFamilyHighAnsi ?? ascii,
    eastAsia: facts.fontFamilyEastAsia ?? paragraph.defaultFontFamilyEastAsia ?? ascii,
    complexScript: facts.fontFamilyCs ?? ascii,
  };
  return Object.freeze({
    fontSizePt,
    fonts: Object.freeze({ ...(facts.fontSlots?.direct ?? fallbackFonts) }),
    themeFonts: facts.fontSlots?.theme ? Object.freeze({ ...facts.fontSlots.theme }) : undefined,
    themeFontPresence: facts.fontSlots?.themePresent
      ? Object.freeze({ ...facts.fontSlots.themePresent }) : undefined,
    weight: (complexScript ? (facts.boldCs ?? false) : (facts.bold ?? false)) ? 700 : 400,
    style: (complexScript ? (facts.italicCs ?? false) : (facts.italic ?? false)) ? 'italic' : 'normal',
    complexScript,
    fontHint: facts.fontHint,
    eastAsiaLanguage: facts.langEastAsia,
    kerning: facts.kerning == null ? undefined : fontSizePt >= facts.kerning,
  });
}

/** Immutable all-run snapshot for the retained paragraph acquisition kernel. */
export function paragraphAcquisitionInput(
  paragraph: DocParagraph,
  source: SourceRef,
): ParagraphAcquisitionInput {
  // Table pagination may have attached legacy cache stamps containing live font
  // resolver functions. They are renderer state, not parser/model facts, and must
  // never cross the retained acquisition boundary.
  const {
    layoutLines: _layoutLines,
    lineSlice: _lineSlice,
    __paragraphTypographyAcquisition: _privateParagraphTypography,
    ...semanticParagraph
  } = paragraph as DocParagraph & Record<string, unknown> & {
    __paragraphTypographyAcquisition?: InternalParagraphTypographyWire;
  };
  const typographyInput = paragraphTypographyAcquisitionInput(paragraph);
  const snapshot = structuredClone(semanticParagraph) as DocParagraph;
  const runs = snapshot.runs.map((run, runIndex): ParagraphAcquisitionRun => {
    if (run.type === 'math') {
      const runRef: SourceRef = Object.freeze({ ...source, path: Object.freeze([...source.path, runIndex]) });
      const internal = run as Partial<InternalMathRun>;
      return Object.freeze({
        ...run,
        source: internal.source ?? runRef,
        resourceKey: internal.resourceKey ?? mathResourceKey(runRef, run.display ? 'display' : 'inline'),
      });
    }
    if (run.type === 'anchorHost') {
      const internal = run as typeof run & { __anchorOccurrenceId?: string };
      const { __anchorOccurrenceId, ...host } = internal;
      return Object.freeze({
        ...host,
        ...(__anchorOccurrenceId === undefined
          ? {}
          : { anchorOccurrenceId: anchorOccurrenceKey(source, __anchorOccurrenceId) }),
      }) as ParagraphAcquisitionRun;
    }
    if (run.type === 'shape' || run.type === 'image' || run.type === 'chart') {
      const originalRun = paragraph.runs[runIndex] as DocRun;
      const localAnchorInput = anchorAcquisitionInput(originalRun);
      const anchorInput = localAnchorInput === undefined
        ? undefined
        : snapshotPlainData({
            ...localAnchorInput,
            occurrenceId: anchorOccurrenceKey(source, localAnchorInput.occurrenceId),
          }, 'DOCX scoped anchor acquisition input');
      const { __anchorAcquisition: _privateAnchor, ...publicRun } = run as typeof run & InternalAnchorRunWire;
      if (run.type !== 'shape') {
        return Object.freeze({
          ...publicRun,
          ...(anchorInput === undefined ? {} : { anchorAcquisitionInput: anchorInput }),
        }) as ParagraphAcquisitionRun;
      }
      const originalShape = originalRun as ShapeRun;
      const vmlTextPathInput = vmlTextPathAcquisitionInput(originalShape);
      const shapeSource: SourceRef = Object.freeze({
        ...source,
        path: Object.freeze([...source.path, runIndex]),
      });
      const textBoxInput = textBoxAcquisitionInput(originalShape, {
        story: 'textbox',
        storyInstance: `${shapeSource.story}:${shapeSource.storyInstance}:${shapeSource.path.join('.')}`,
        path: [],
      });
      return Object.freeze({
        ...publicRun,
        ...(vmlTextPathInput === undefined ? {} : { vmlTextPathInput }),
        ...(textBoxInput.length === 0 ? {} : { textBoxInput }),
        ...(anchorInput === undefined ? {} : { anchorAcquisitionInput: anchorInput }),
      }) as ParagraphAcquisitionRun;
    }
    if (run.type === 'text' || run.type === 'field') {
      const originalRun = paragraph.runs[runIndex] as Extract<DocRun, { type: 'text' | 'field' }>;
      const runTypographyInput = runTypographyAcquisitionInput(originalRun);
      const {
        __typographyAcquisition: _privateRunTypography,
        ...publicRun
      } = run as typeof run & { __typographyAcquisition?: InternalRunTypographyWire };
      return Object.freeze({
        ...publicRun,
        ...(runTypographyInput === undefined ? {} : { typographyInput: runTypographyInput }),
      }) as ParagraphAcquisitionRun;
    }
    return Object.freeze({ ...run }) as ParagraphAcquisitionRun;
  });
  return deepFreezePlainData({
    ...snapshot,
    runs: runs as ParagraphAcquisitionRun[],
    numberingMarkerShapeInput: paragraph.numbering
      ? numberingMarkerShapeInput(
          paragraph.numbering,
          paragraph.runs.find(
            (run): run is Extract<DocRun, { type: 'text' | 'field' }> =>
              run.type === 'text' || run.type === 'field',
          )?.fontSize ?? paragraph.defaultFontSize ?? 10,
        )
      : undefined,
    paragraphMarkShapeInput: paragraphMarkShapeInput(paragraph),
    ...(typographyInput === undefined ? {} : { typographyInput }),
  }) as unknown as ParagraphAcquisitionInput;
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
  const document = (changed
    ? { ...doc, body, headers, footers, footnotes, endnotes }
    : doc) as InternalDocxDocumentModel;
  const sectionPlacementInputs = projectSectionPlacementInputs(document);
  sectionPlacementInputsByDocument.set(document, sectionPlacementInputs);
  setBodySectionPlacementInputs(document.body, document.section, sectionPlacementInputs);
  return Object.freeze({
    document,
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
