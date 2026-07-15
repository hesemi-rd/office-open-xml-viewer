import type { DocParagraph, ParagraphBorders, ParaBorderEdge } from '../types.js';

export type ParagraphBorderEdges = Readonly<{
  top: 'top' | 'between' | 'none';
  bottom: 'bottom' | 'none';
}>;

function effectiveEdge(edge: ParaBorderEdge | null): ParaBorderEdge | null {
  return edge == null || edge.style === 'none' ? null : edge;
}

function sameParagraphEdge(a: ParaBorderEdge | null, b: ParaBorderEdge | null): boolean {
  const left = effectiveEdge(a);
  const right = effectiveEdge(b);
  if (left == null || right == null) return left == null && right == null;
  return left.style === right.style
    && left.width === right.width
    && (left.space ?? 0) === (right.space ?? 0)
    && (left.color ?? null) === (right.color ?? null);
}

function sameParagraphBorders(
  a: ParagraphBorders | null | undefined,
  b: ParagraphBorders | null | undefined,
): boolean {
  if (!a || !b) return false;
  return sameParagraphEdge(a.top, b.top)
    && sameParagraphEdge(a.bottom, b.bottom)
    && sameParagraphEdge(a.left, b.left)
    && sameParagraphEdge(a.right, b.right)
    && sameParagraphEdge(a.between, b.between);
}

export function hasVisibleParagraphBorder(
  borders: ParagraphBorders | null | undefined,
): boolean {
  if (!borders) return false;
  return [borders.top, borders.right, borders.bottom, borders.left, borders.between]
    .some((edge) => edge != null && edge.style !== 'none');
}

/** Pure §17.3.1.7 matching predicate; callers supply actual flow adjacency. */
export function paragraphsShareBorderBox(
  previous: DocParagraph | null,
  current: DocParagraph | null,
): boolean {
  if (!previous || !current || previous.framePr || current.framePr) return false;
  return hasVisibleParagraphBorder(previous.borders)
    && hasVisibleParagraphBorder(current.borders)
    && sameParagraphBorders(previous.borders, current.borders);
}

/** Resolves edge ownership once for any retained paragraph flow container. */
export function resolveParagraphBorderEdges(
  previous: DocParagraph | null,
  current: DocParagraph,
  next: DocParagraph | null,
  groupedFrameFlow = false,
): ParagraphBorderEdges {
  const shares = (left: DocParagraph | null, right: DocParagraph | null): boolean =>
    groupedFrameFlow
      ? !!left && !!right
        && !!left.framePr && !!right.framePr
        && hasVisibleParagraphBorder(left.borders)
        && hasVisibleParagraphBorder(right.borders)
        && sameParagraphBorders(left.borders, right.borders)
      : paragraphsShareBorderBox(left, right);
  const joinsPrevious = shares(previous, current);
  const joinsNext = shares(current, next);
  const between = current.borders?.between;
  return Object.freeze({
    top: joinsPrevious
      ? between && between.style !== 'none' ? 'between' : 'none'
      : 'top',
    bottom: joinsNext ? 'none' : 'bottom',
  });
}
