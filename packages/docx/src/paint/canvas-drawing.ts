import type { DrawingLayout } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';
import { canvasFontString } from '@silurus/ooxml-core';

export function paintDrawingLayout(node: DrawingLayout, context: CanvasPaintContext): void {
  for (const command of node.commands) {
    if (command.kind === 'fill-rect') {
      context.ctx.fillStyle = command.fill;
      context.ctx.fillRect(
        command.rect.xPt,
        command.rect.yPt,
        command.rect.widthPt,
        command.rect.heightPt,
      );
      continue;
    }
    if (command.kind === 'stroke-rect') {
      context.ctx.strokeStyle = command.stroke;
      context.ctx.lineWidth = command.lineWidthPt;
      context.ctx.setLineDash([...command.dashPt]);
      context.ctx.strokeRect(
        command.rect.xPt,
        command.rect.yPt,
        command.rect.widthPt,
        command.rect.heightPt,
      );
      context.ctx.setLineDash([]);
      continue;
    }
    context.ctx.fillStyle = command.fill;
    context.ctx.font = canvasFontString(
      command.fontRoute,
      command.fontSizePt,
      command.fontWeight,
      command.fontStyle,
    );
    context.ctx.textAlign = command.align === 'start' ? 'left' : command.align === 'end' ? 'right' : 'center';
    context.ctx.textBaseline = command.baseline;
    const xPt = command.align === 'start'
      ? command.rect.xPt
      : command.align === 'end'
        ? command.rect.xPt + command.rect.widthPt
        : command.rect.xPt + command.rect.widthPt / 2;
    const yPt = command.baseline === 'top'
      ? command.rect.yPt
      : command.baseline === 'bottom'
        ? command.rect.yPt + command.rect.heightPt
        : command.rect.yPt + command.rect.heightPt / 2;
    context.ctx.fillText(command.text, xPt, yPt);
  }
}
