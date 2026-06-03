import type { MathBox, DrawOp } from './layout';
import type { MathStyle } from '../types/math';

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

function styleToFont(style: MathStyle, sizePx: number, family: string): string {
  const bold = style === 'bold' || style === 'boldItalic';
  const italic = style === 'italic' || style === 'boldItalic';
  return `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px "${family}"`;
}

/**
 * Draw a laid-out math box. `(x, baseline)` is the box origin in canvas px
 * (baseline y measured from the canvas top, matching `ctx.fillText`).
 */
export function renderMathBox(
  ctx: Ctx2D,
  box: MathBox,
  x: number,
  baseline: number,
  color: string,
  family: string,
): void {
  ctx.save();
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  for (const op of box.ops) drawOp(ctx, op, x, baseline, family);
  ctx.restore();
}

function drawOp(ctx: Ctx2D, op: DrawOp, x: number, baseline: number, family: string): void {
  if (op.type === 'glyph') {
    ctx.font = styleToFont(op.style, op.sizePx, family);
    ctx.fillText(op.text, x + op.x, baseline + op.y);
  } else if (op.type === 'rule') {
    ctx.fillRect(x + op.x, baseline + op.y, op.w, op.h);
  } else {
    // Polyline stroke (synthesized radical sign).
    ctx.save();
    ctx.lineWidth = op.lineWidth;
    ctx.lineJoin = 'miter';
    ctx.lineCap = 'butt';
    ctx.beginPath();
    op.points.forEach((p, i) => {
      const px = x + p.x;
      const py = baseline + p.y;
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    });
    ctx.stroke();
    ctx.restore();
  }
}

export interface MathMetrics {
  width: number;
  ascent: number;
  descent: number;
}

export function measureMathBox(box: MathBox): MathMetrics {
  return { width: box.width, ascent: box.ascent, descent: box.descent };
}
