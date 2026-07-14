import type { DrawingLayout } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';

export function paintDrawingLayout(node: DrawingLayout, context: CanvasPaintContext): void {
  for (const command of node.commands) {
    context.ctx.fillStyle = command.fill;
    context.ctx.fillRect(
      command.rect.xPt,
      command.rect.yPt,
      command.rect.widthPt,
      command.rect.heightPt,
    );
  }
}
