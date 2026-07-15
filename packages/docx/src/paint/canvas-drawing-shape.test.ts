import { describe, expect, it } from 'vitest';
import type { DrawingLayout } from '../layout/types.js';
import { paintDrawingLayout } from './canvas-drawing.js';
import type { CanvasPaintContext } from './types.js';

function shapeDrawing(): DrawingLayout {
  const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
  return {
    kind: 'drawing',
    id: 'shape-1',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body',
    flowBounds: bounds,
    inkBounds: bounds,
    advancePt: 50,
    ordinaryFlow: false,
    commands: [{
      kind: 'drawingml-shape',
      plan: {
        rect: { x: 10, y: 20, w: 100, h: 50 },
        geometry: { kind: 'preset', name: 'rect', adjustments: [] },
        fill: { fillType: 'solid', color: 'FF0000' },
        stroke: { color: '000000', width: 2 },
        transform: { rotationDeg: 0, flipH: false, flipV: false },
      },
    }],
  };
}

function recordingContext(): { context: CanvasPaintContext; operations: string[] } {
  const operations: string[] = [];
  const ctx = {
    fillStyle: '', strokeStyle: '', lineWidth: 1, lineCap: 'butt',
    globalAlpha: 1, font: '', textAlign: 'left', textBaseline: 'alphabetic',
    save: () => operations.push('save'),
    restore: () => operations.push('restore'),
    beginPath: () => operations.push('beginPath'),
    rect: () => operations.push('rect'),
    fill: () => operations.push('fill'),
    stroke: () => operations.push('stroke'),
    setLineDash: () => operations.push('setLineDash'),
    translate: () => operations.push('translate'),
    rotate: () => operations.push('rotate'),
    scale: () => operations.push('scale'),
    createLinearGradient: () => {
      operations.push('createLinearGradient');
      return { addColorStop: () => operations.push('addColorStop') };
    },
    fillText: () => operations.push('fillText'),
    measureText: () => { throw new Error('retained drawing paint must not measure'); },
  } as unknown as CanvasRenderingContext2D;
  return {
    operations,
    context: {
      ctx,
      scale: 4,
      dpr: 2,
      resources: { paint: () => { throw new Error('shape must not use the resource painter'); } },
    },
  };
}

describe('retained DrawingML shape painting', () => {
  it('dispatches explicit shape plans to the shared point-space painter', () => {
    const { context, operations } = recordingContext();

    paintDrawingLayout(shapeDrawing(), context);

    expect(operations).toEqual(expect.arrayContaining([
      'save', 'beginPath', 'rect', 'fill', 'setLineDash', 'stroke', 'restore',
    ]));
  });

  it('paints a shaped watermark command without measuring', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
    const drawing: DrawingLayout = {
      kind: 'drawing', id: 'watermark',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, ordinaryFlow: false,
      commands: [{
        kind: 'watermark-text', rect: bounds, text: 'DRAFT',
        fill: { fillType: 'solid', color: '808080' },
        opacity: .4, rotationDeg: 315, fitShape: true, fontSizePt: 36,
        sourceBounds: { xPt: -5, yPt: -80, widthPt: 250, heightPt: 100 },
        spans: [{
          text: 'DRAFT', advancePt: 250,
          fontRoute: { familyList: 'Arial', scope: 'native', fingerprint: 'arial' },
          fontWeight: 700, fontStyle: 'italic',
        }],
      }],
    };

    paintDrawingLayout(drawing, context);

    expect(operations).toEqual([
      'save', 'translate', 'rotate', 'scale', 'translate', 'fillText', 'restore',
    ]);
  });

  it('retains gradient fill semantics and skips no-fill text without measuring', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 };
    const command = {
      kind: 'watermark-text' as const, rect: bounds, text: 'DRAFT',
      opacity: 1, rotationDeg: 0, fitShape: false, fontSizePt: 12,
      sourceBounds: { xPt: 0, yPt: -8, widthPt: 30, heightPt: 10 },
      spans: [{
        text: 'DRAFT', advancePt: 30,
        fontRoute: { familyList: 'Arial', scope: 'native' as const, fingerprint: 'arial' },
        fontWeight: 400, fontStyle: 'normal' as const,
      }],
    };
    const layout = (fill: null | { fillType: 'gradient'; angle: number; gradType: string; stops: { position: number; color: string }[] }): DrawingLayout => ({
      kind: 'drawing', id: 'watermark-fill', source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds, advancePt: 0, ordinaryFlow: false,
      commands: [{ ...command, fill }],
    });
    paintDrawingLayout(layout(null), context);
    expect(operations).toEqual([]);

    paintDrawingLayout(layout({
      fillType: 'gradient', angle: 0, gradType: 'linear',
      stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
    }), context);
    expect(operations).toEqual([
      'createLinearGradient', 'addColorStop', 'addColorStop',
      'save', 'translate', 'translate', 'fillText', 'restore',
    ]);
  });

  it('ignores explicit no-op drawing commands', () => {
    const { context, operations } = recordingContext();
    const bounds = { xPt: 0, yPt: 0, widthPt: 1, heightPt: 1 };
    paintDrawingLayout({
      kind: 'drawing', id: 'noop', source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds,
      advancePt: 0, ordinaryFlow: false, commands: [{ kind: 'noop' }],
    }, context);
    expect(operations).toEqual([]);
  });
});
