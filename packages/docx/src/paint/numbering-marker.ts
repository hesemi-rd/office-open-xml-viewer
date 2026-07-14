import { canvasFontString } from '@silurus/ooxml-core';

type MarkerSpan = Readonly<{
  text: string;
  advancePt: number;
  fontRoute: Parameters<typeof canvasFontString>[0];
  font: Readonly<{ weight: number; style: 'normal' | 'italic' }>;
}>;

export interface NumberingMarkerPaintLayout {
  readonly shape: Readonly<{ spans: readonly MarkerSpan[] }>;
  readonly fontSizePx: number;
}

export type VerticalNumberingMarkerPainter = (
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  text: string,
  x: number,
  baseline: number,
  fontSizePx: number,
) => void;

/** Serialize exactly the routed spans that supplied marker geometry. */
export function paintNumberingMarkerText(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  layout: NumberingMarkerPaintLayout,
  x: number,
  baseline: number,
  verticalPainter?: VerticalNumberingMarkerPainter,
): void {
  let advancePx = 0;
  for (const span of layout.shape.spans) {
    ctx.font = canvasFontString(
      span.fontRoute,
      layout.fontSizePx,
      span.font.weight,
      span.font.style,
    );
    if (verticalPainter) {
      verticalPainter(ctx, span.text, x + advancePx, baseline, layout.fontSizePx);
    } else {
      ctx.fillText(span.text, x + advancePx, baseline);
    }
    advancePx += span.advancePt;
  }
}
