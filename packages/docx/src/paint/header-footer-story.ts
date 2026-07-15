import {
  withDeferredFrontPaintSession,
  type DeferredFrontPaintState,
} from './deferred-front-session.js';

export interface HeaderFooterBorderMerge {
  suppressTop: boolean;
  suppressBottom: boolean;
}

export interface HeaderFooterStoryState extends DeferredFrontPaintState {
  y: number;
  scale: number;
}

export interface HeaderFooterStoryElement {
  readonly type: string;
}

export interface HeaderFooterStory<Element extends HeaderFooterStoryElement> {
  readonly body: readonly Element[];
}

export interface HeaderFooterStoryOperations<
  State extends HeaderFooterStoryState,
  Element extends HeaderFooterStoryElement,
  Paragraph,
  Table,
  Borders,
> {
  preRegisterPageFloats(elements: readonly Element[], state: State): void;
  paragraphOf(element: Element): Paragraph;
  tableOf(element: Element): Table;
  hasFrame(paragraph: Paragraph): boolean;
  frameAnchorLineHeight(
    elements: readonly Element[],
    element: Element,
    state: State,
  ): number;
  paintFrameParagraph(paragraph: Paragraph, state: State, lineHeight: number): void;
  spaceBefore(paragraph: Paragraph): number;
  spaceAfter(paragraph: Paragraph): number;
  bordersOf(paragraph: Paragraph): Borders | null | undefined;
  contextualSpacing(
    previous: Paragraph | null,
    paragraph: Paragraph,
    previousSpaceAfter: number,
    spaceBefore: number,
  ): Readonly<{ suppressBefore: boolean; overlap: number }>;
  hasBorder(borders: Borders | null | undefined): boolean;
  sharesBorder(previous: Paragraph | null, paragraph: Paragraph | null): boolean;
  paintParagraph(
    paragraph: Paragraph,
    state: State,
    suppressBefore: boolean,
    borderMerge: HeaderFooterBorderMerge | undefined,
  ): void;
  paintTable(table: Table, state: State): void;
  tableResetsParagraphFlow(table: Table): boolean;
}

export type HeaderFooterStoryPainter<
  State extends HeaderFooterStoryState,
  Element extends HeaderFooterStoryElement,
> = (
  headerFooter: HeaderFooterStory<Element>,
  topY: number,
  base: State,
) => number;

/**
 * Owns header/footer story traversal while Series A migrates only the body.
 * B1 removes this adapter when every story is acquired into the same retained
 * graph; B3 then replaces its deferred-front session with immutable PageLayers.
 */
export function createHeaderFooterStoryPainter<
  State extends HeaderFooterStoryState,
  Element extends HeaderFooterStoryElement,
  Paragraph,
  Table,
  Borders,
>(
  operations: HeaderFooterStoryOperations<State, Element, Paragraph, Table, Borders>,
): HeaderFooterStoryPainter<State, Element> {
  return (headerFooter, topY, base) => {
    const state: State = { ...base, y: topY };
    const elements = headerFooter.body;
    let previousParagraph: Paragraph | null = null;
    let previousSpaceAfter = 0;
    operations.preRegisterPageFloats(elements, state);
    withDeferredFrontPaintSession(state, () => {
      for (let index = 0; index < elements.length; index++) {
        const element = elements[index];
        if (element.type === 'paragraph') {
          const paragraph = operations.paragraphOf(element);
          if (operations.hasFrame(paragraph)) {
            operations.paintFrameParagraph(
              paragraph,
              state,
              operations.frameAnchorLineHeight(elements, element, state),
            );
            continue;
          }
          const adjust = operations.contextualSpacing(
            previousParagraph,
            paragraph,
            previousSpaceAfter,
            operations.spaceBefore(paragraph),
          );
          state.y -= adjust.overlap * state.scale;
          const previous = elements[index - 1]?.type === 'paragraph'
            ? operations.paragraphOf(elements[index - 1])
            : null;
          const next = elements[index + 1]?.type === 'paragraph'
            ? operations.paragraphOf(elements[index + 1])
            : null;
          const borderMerge = operations.hasBorder(operations.bordersOf(paragraph))
            ? {
                suppressTop: operations.sharesBorder(previous, paragraph),
                suppressBottom: operations.sharesBorder(paragraph, next),
              }
            : undefined;
          operations.paintParagraph(
            paragraph,
            state,
            adjust.suppressBefore,
            borderMerge,
          );
          previousParagraph = paragraph;
          previousSpaceAfter = operations.spaceAfter(paragraph);
          continue;
        }
        if (element.type === 'table') {
          const table = operations.tableOf(element);
          operations.paintTable(table, state);
          if (operations.tableResetsParagraphFlow(table)) {
            previousParagraph = null;
            previousSpaceAfter = 0;
          }
        }
      }
    });
    return state.y;
  };
}
