import { describe, expect, it, vi } from 'vitest';
import { applyFrameTransform } from './renderer.js';
import type { ChartElement, MediaElement, TableElement } from './types.js';

function contextRecorder() {
  const translate = vi.fn();
  const rotate = vi.fn();
  const scale = vi.fn();
  return {
    ctx: { translate, rotate, scale } as unknown as CanvasRenderingContext2D,
    translate,
    rotate,
    scale,
  };
}

const frame = {
  x: 10,
  y: 20,
  width: 40,
  height: 10,
  rotation: 90,
  flipH: true,
  flipV: true,
};

describe('PPTX frame transforms (ECMA-376 Annex L)', () => {
  it.each([
    { type: 'table', cols: [], rows: [] } as unknown as TableElement,
    { type: 'chart', chart: {} } as unknown as ChartElement,
    {
      type: 'media',
      mediaKind: 'video',
      posterPath: '',
      posterMimeType: '',
      mediaPath: '',
      mimeType: 'video/mp4',
    } as unknown as MediaElement,
  ])('applies rotation and both flips to a $type frame', (element) => {
    const { ctx, translate, rotate, scale } = contextRecorder();
    applyFrameTransform(ctx, { ...element, ...frame }, 2);

    expect(translate.mock.calls).toEqual([
      [60, 50],
      [-60, -50],
    ]);
    expect(rotate).toHaveBeenCalledWith(Math.PI / 2);
    expect(scale.mock.calls).toEqual([
      [-1, 1],
      [1, -1],
    ]);
  });
});
