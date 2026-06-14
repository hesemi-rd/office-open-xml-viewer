import type { ArrowEnd, Stroke } from '../types/common';
import { hexToRgba } from './paint';

/**
 * Draw a DrawingML line-end decoration (arrow head) at `(tipX, tipY)`,
 * oriented along `angle` radians (0 = pointing right, +x axis).
 *
 * ECMA-376 §20.1.8.3 (CT_LineEndProperties) / §20.1.10.33 (ST_LineEndType:
 * none / triangle / stealth / diamond / oval / arrow) / §20.1.10.31–.32
 * (ST_LineEndWidth / ST_LineEndLength: sm / med / lg). The spec only names
 * the w/len steps as *relative* sizes, not exact ratios — the multiples of
 * line width below are calibrated against PowerPoint's rendering and shared
 * between the pptx and docx renderers so connector arrows look identical.
 *
 * `scale` is the EMU → device-px factor (same convention as core's
 * `applyStroke`, where stroke width in px is `stroke.width * scale`).
 */
export function drawArrowHead(
  ctx: CanvasRenderingContext2D,
  tipX: number,
  tipY: number,
  angle: number,
  arrowEnd: ArrowEnd,
  stroke: Stroke,
  scale: number,
): void {
  if (arrowEnd.type === 'none') return;
  const lw = Math.max(0.5, stroke.width * scale);
  const wMul = arrowEnd.w === 'sm' ? 4 : arrowEnd.w === 'lg' ? 8 : 6;
  const lMul = arrowEnd.len === 'sm' ? 4 : arrowEnd.len === 'lg' ? 8 : 6;
  const halfW = (lw * wMul) / 2;
  const len = lw * lMul;
  const color = hexToRgba(stroke.color);

  ctx.save();
  ctx.translate(tipX, tipY);
  ctx.rotate(angle);
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  ctx.lineWidth = lw;
  ctx.setLineDash([]);
  ctx.beginPath();
  switch (arrowEnd.type) {
    case 'triangle':
    case 'stealth':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len, -halfW);
      ctx.lineTo(-len, halfW);
      ctx.closePath();
      ctx.fill();
      break;
    case 'arrow':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len, -halfW);
      ctx.moveTo(0, 0);
      ctx.lineTo(-len, halfW);
      ctx.stroke();
      break;
    case 'diamond':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len / 2, -halfW);
      ctx.lineTo(-len, 0);
      ctx.lineTo(-len / 2, halfW);
      ctx.closePath();
      ctx.fill();
      break;
    case 'oval':
      ctx.ellipse(-len / 2, 0, len / 2, halfW, 0, 0, Math.PI * 2);
      ctx.fill();
      break;
  }
  ctx.restore();
}
