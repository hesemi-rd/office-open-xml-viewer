import { docxBorderDashArray } from '@silurus/ooxml-core';
import type { BorderSegment } from './types.js';

/** Converts an authored ST_Border token to clone-safe, paint-ready treatment. */
export function retainedBorderTreatment(
  authoredStyle: string,
  widthPt: number,
): Pick<BorderSegment, 'authoredStyle' | 'style' | 'dashPatternPt'> {
  const dashPatternPt = docxBorderDashArray(authoredStyle, widthPt);
  return Object.freeze({
    authoredStyle,
    style: authoredStyle === 'double'
      ? 'double'
      : dashPatternPt.length > 0
        ? 'dashed'
        : authoredStyle.includes('wave') ? 'wavy' : 'solid',
    dashPatternPt: Object.freeze(dashPatternPt),
  });
}
