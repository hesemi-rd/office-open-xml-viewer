import { describe, expect, it } from 'vitest';
import {
  paintDrawingMLShape,
  type DrawingMLShapePaintPlan,
} from './drawingml-shape.js';

function recordingContext() {
  const operations: Array<{ name: string; args: unknown[] }> = [];
  let lineWidth = 1;
  let lineCap: CanvasLineCap = 'butt';
  const gradient = {
    addColorStop(...args: unknown[]) { operations.push({ name: 'addColorStop', args }); },
  } as unknown as CanvasGradient;
  const ctx = {
    fillStyle: '', strokeStyle: '',
    get lineWidth() { return lineWidth; },
    set lineWidth(value: number) { lineWidth = value; operations.push({ name: 'lineWidth', args: [value] }); },
    get lineCap() { return lineCap; },
    set lineCap(value: CanvasLineCap) { lineCap = value; operations.push({ name: 'lineCap', args: [value] }); },
    save() { operations.push({ name: 'save', args: [] }); },
    restore() { operations.push({ name: 'restore', args: [] }); },
    translate(...args: unknown[]) { operations.push({ name: 'translate', args }); },
    rotate(...args: unknown[]) { operations.push({ name: 'rotate', args }); },
    scale(...args: unknown[]) { operations.push({ name: 'scale', args }); },
    beginPath() { operations.push({ name: 'beginPath', args: [] }); },
    rect(...args: unknown[]) { operations.push({ name: 'rect', args }); },
    moveTo(...args: unknown[]) { operations.push({ name: 'moveTo', args }); },
    lineTo(...args: unknown[]) { operations.push({ name: 'lineTo', args }); },
    bezierCurveTo(...args: unknown[]) { operations.push({ name: 'bezierCurveTo', args }); },
    ellipse(...args: unknown[]) { operations.push({ name: 'ellipse', args }); },
    closePath() { operations.push({ name: 'closePath', args: [] }); },
    fill(...args: unknown[]) { operations.push({ name: 'fill', args }); },
    stroke() { operations.push({ name: 'stroke', args: [] }); },
    setLineDash(...args: unknown[]) { operations.push({ name: 'setLineDash', args }); },
    createLinearGradient(...args: unknown[]) {
      operations.push({ name: 'createLinearGradient', args }); return gradient;
    },
    createRadialGradient(...args: unknown[]) {
      operations.push({ name: 'createRadialGradient', args }); return gradient;
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, operations };
}

describe('shared DrawingML shape painter', () => {
  it('renders preset geometry, gradients, transforms, and point-width strokes once', () => {
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 10, y: 20, w: 100, h: 50 },
      geometry: { kind: 'preset', name: 'rect', adjustments: [] },
      fill: {
        fillType: 'gradient', angle: 0, gradType: 'linear',
        stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: '0000FF' }],
      },
      stroke: { color: '000000', width: 2, dashStyle: 'dash', lineCap: 'round' },
      transform: { rotationDeg: 90, flipH: true, flipV: false },
    };
    const { ctx, operations } = recordingContext();

    paintDrawingMLShape(ctx, plan, 1);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'translate', args: [60, 45] },
      { name: 'rotate', args: [Math.PI / 2] },
      { name: 'scale', args: [-1, 1] },
      { name: 'createLinearGradient', args: [10, 45, 110, 45] },
      { name: 'addColorStop', args: [0, 'rgba(255,0,0,1)'] },
      { name: 'addColorStop', args: [1, 'rgba(0,0,255,1)'] },
      { name: 'lineWidth', args: [2] },
      { name: 'lineCap', args: ['round'] },
    ]));
  });

  it('renders custom cubic geometry and its terminal arrows', () => {
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 10, y: 20, w: 100, h: 50 },
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'cubicBezTo', x1: .25, y1: 0, x2: .75, y2: 1, x: 1, y: 1 },
      ]] },
      fill: null,
      stroke: {
        color: '123456', width: 1, lineCap: 'square',
        headEnd: { type: 'arrow', w: 'sm', len: 'sm' },
        tailEnd: { type: 'triangle', w: 'med', len: 'lg' },
      },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    };
    const { ctx, operations } = recordingContext();

    paintDrawingMLShape(ctx, plan, 1);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'moveTo', args: [10, 20] },
      { name: 'bezierCurveTo', args: [35, 20, 85, 70, 110, 70] },
      { name: 'translate', args: [10, 20] },
      { name: 'translate', args: [110, 70] },
    ]));
  });

  it('preserves arbitrary adjustment counts for preset geometry', () => {
    const adjustments = Array.from({ length: 12 }, (_, index) => index * 1000);
    const plan: DrawingMLShapePaintPlan = {
      rect: { x: 0, y: 0, w: 100, h: 100 },
      geometry: { kind: 'preset', name: 'rect', adjustments },
      fill: { fillType: 'solid', color: 'FFFFFF' }, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    };
    const { ctx } = recordingContext();

    expect(() => paintDrawingMLShape(ctx, plan, 1)).not.toThrow();
    expect(plan.geometry.kind === 'preset' && plan.geometry.adjustments).toHaveLength(12);
  });
});
