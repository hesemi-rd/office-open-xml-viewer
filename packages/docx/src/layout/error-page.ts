import type { SectionLayoutContext } from '../layout-context.js';
import type { TextLayoutService } from './text.js';
import type {
  DocumentLayout,
  DrawingLayout,
  DrawingPaintCommand,
  LayoutRect,
  SourceRef,
} from './types.js';

export interface ErrorPageSize {
  readonly widthPt: number;
  readonly heightPt: number;
}

const ERROR_SOURCE: SourceRef = Object.freeze({
  story: 'body',
  storyInstance: 'parse-error',
  path: Object.freeze([]),
});

function wrapWords(
  text: string,
  maxWidthPt: number,
  fontSizePt: number,
  service: TextLayoutService,
  maxLines: number,
): string[] {
  const words = text.trim().split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let line = '';
  for (const word of words) {
    const candidate = line ? `${line} ${word}` : word;
    const shaped = service.shape({
      text: candidate,
      fontSizePt,
      fonts: { ascii: 'sans-serif', highAnsi: 'sans-serif', eastAsia: 'sans-serif', complexScript: 'sans-serif' },
      genericFamily: 'sans-serif',
    });
    if (line && shaped.advancePt > maxWidthPt) {
      lines.push(line);
      line = word;
      if (lines.length === maxLines) break;
    } else {
      line = candidate;
    }
  }
  if (line && lines.length < maxLines) lines.push(line);
  return lines;
}

export function layoutParseErrorPage(
  message: string,
  size: ErrorPageSize,
  text: TextLayoutService,
): DocumentLayout {
  if (!(size.widthPt > 0 && size.heightPt > 0)) throw new RangeError('Error page size must be positive');
  const padPt = Math.max(18, Math.min(size.widthPt, size.heightPt) * 0.06);
  const frame: LayoutRect = {
    xPt: padPt,
    yPt: padPt,
    widthPt: size.widthPt - padPt * 2,
    heightPt: size.heightPt - padPt * 2,
  };
  const detailSizePt = Math.max(8, Math.min(size.widthPt, size.heightPt) * 0.02);
  const detailLines = wrapWords(message, size.widthPt - padPt * 4, detailSizePt, text, 4);
  const detailLineHeightPt = detailSizePt * 1.4;
  const commands: DrawingPaintCommand[] = [
    { kind: 'fill-rect', rect: { xPt: 0, yPt: 0, widthPt: size.widthPt, heightPt: size.heightPt }, fill: '#ffffff' },
    { kind: 'stroke-rect', rect: frame, stroke: '#c8ccd2', lineWidthPt: 1, dashPt: [6, 5] },
    {
      kind: 'text', rect: { xPt: 0, yPt: size.heightPt * 0.27, widthPt: size.widthPt, heightPt: 36 },
      text: '⚠', fill: '#b23b3b', fontFamily: 'sans-serif', fontSizePt: 28,
      fontWeight: 400, fontStyle: 'normal', align: 'center', baseline: 'middle',
    },
    {
      kind: 'text', rect: { xPt: padPt * 2, yPt: size.heightPt * 0.40, widthPt: size.widthPt - padPt * 4, heightPt: 24 },
      text: 'This document could not be displayed', fill: '#333333', fontFamily: 'sans-serif', fontSizePt: 13,
      fontWeight: 600, fontStyle: 'normal', align: 'center', baseline: 'middle',
    },
    ...detailLines.map((line, index): DrawingPaintCommand => ({
      kind: 'text',
      rect: {
        xPt: padPt * 2,
        yPt: size.heightPt * 0.50 + detailLineHeightPt * index,
        widthPt: size.widthPt - padPt * 4,
        heightPt: detailLineHeightPt,
      },
      text: line,
      fill: '#666666',
      fontFamily: 'sans-serif',
      fontSizePt: detailSizePt,
      fontWeight: 400,
      fontStyle: 'normal',
      align: 'center',
      baseline: 'middle',
    })),
  ];
  const node: DrawingLayout = {
    kind: 'drawing',
    id: 'parse-error-page',
    source: ERROR_SOURCE,
    flowDomainId: 'parse-error',
    flowBounds: frame,
    inkBounds: frame,
    advancePt: frame.heightPt,
    ordinaryFlow: false,
    commands,
  };
  const section: SectionLayoutContext = {
    geometry: {
      pageWidth: size.widthPt,
      pageHeight: size.heightPt,
      marginTop: padPt,
      marginRight: padPt,
      marginBottom: padPt,
      marginLeft: padPt,
      headerDistance: 0,
      footerDistance: 0,
    },
    columns: [{ xPt: padPt, wPt: frame.widthPt }],
    grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
    textDirection: 'lrTb',
    verticalAlignment: 'top',
  };
  return {
    pages: [{
      pageIndex: 0,
      geometry: {
        xPt: 0,
        yPt: 0,
        widthPt: size.widthPt,
        heightPt: size.heightPt,
        contentTopPt: padPt,
        contentBottomPt: size.heightPt - padPt,
      },
      flowDomains: [{ id: 'parse-error', kind: 'body', bounds: frame }],
      section,
      layers: {
        paintOrder: [{ layer: 'body', nodeId: node.id }],
        background: [], behindText: [], header: [], body: [node], notes: [], front: [], footer: [],
      },
      readingOrder: [node.id],
    }],
    diagnostics: [{
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      source: ERROR_SOURCE,
      message,
    }],
  };
}
