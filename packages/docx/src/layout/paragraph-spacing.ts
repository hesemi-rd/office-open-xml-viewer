export interface ParagraphSpacingParticipant {
  readonly contextualSpacing?: boolean;
  readonly styleId?: string | null;
}

/**
 * ECMA-376 §17.3.1.33 paragraph gap, with the per-side
 * `w:contextualSpacing` suppression from §17.3.1.9.
 *
 * The returned value is the complete distance between the preceding line block
 * and the current line block. Callers use this one rule for ordinary flow,
 * table-cell block folds, and DrawingML/WPS text boxes.
 */
export function paragraphGapPt(
  previous: ParagraphSpacingParticipant | null,
  current: ParagraphSpacingParticipant,
  previousAfterPt: number,
  currentBeforePt: number,
): number {
  if (!previous) return currentBeforePt;
  const sameStyle = !!(previous.styleId && previous.styleId === current.styleId);
  const dropPrevious = !!(sameStyle && previous.contextualSpacing);
  const dropCurrent = !!(sameStyle && current.contextualSpacing);
  if (dropPrevious && dropCurrent) return 0;
  if (dropCurrent) return previousAfterPt;
  if (dropPrevious) return Math.max(currentBeforePt - previousAfterPt, 0);
  return Math.max(previousAfterPt, currentBeforePt);
}

/** Legacy cursor adjustment expressed from the shared total-gap authority. */
export function paragraphGapAdjustment(
  previous: ParagraphSpacingParticipant | null,
  current: ParagraphSpacingParticipant,
  previousAfterPt: number,
  currentBeforePt: number,
): { readonly suppressBefore: boolean; readonly overlap: number } {
  const gapPt = paragraphGapPt(previous, current, previousAfterPt, currentBeforePt);
  const suppressBefore = gapPt <= previousAfterPt;
  return {
    suppressBefore,
    overlap: previousAfterPt + (suppressBefore ? 0 : currentBeforePt) - gapPt,
  };
}
