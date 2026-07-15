import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { stableFingerprint } from './fingerprint.js';
import { createPaintResourceRegistry } from './paint-resources.js';
import type { PaintResourceDescriptor } from './types.js';

function chartModel(title: string): ChartModel {
  return {
    chartType: 'line',
    title,
    categories: ['A'],
    series: [{ name: 'Series', values: [1], color: null }],
    varyColors: false,
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: false,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'cross',
    catAxisMajorTickMark: 'cross',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: null,
    valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
  } as ChartModel;
}

function descriptors(): PaintResourceDescriptor[] {
  return [{
    kind: 'math',
    resourceKey: 'math:body:0:inline',
  }, {
    kind: 'image',
    resourceKey: 'image:body:0:word%2Fmedia%2Fimage1.png',
    partPath: 'word/media/image1.png',
    mimeType: 'image/png',
    intrinsicSize: { widthPt: 40, heightPt: 30 },
    svgImagePath: 'word/media/image1.svg',
    srcRect: { l: 0.1, t: 0.2, r: 0.1, b: 0 },
    rotation: 15,
    flipH: true,
    flipV: true,
    alpha: 0.5,
    colorReplaceFrom: 'FFFFFF',
    duotone: { clr1: '000000', clr2: 'FFFFFF' },
  }, {
    kind: 'chart',
    resourceKey: 'chart:body:0',
    intrinsicSize: { widthPt: 120, heightPt: 80 },
    model: chartModel('Original'),
  }, {
    kind: 'picture-bullet',
    resourceKey: 'image:body:1:word%2Fmedia%2Fbullet.gif',
    partPath: 'word/media/bullet.gif',
    mimeType: 'image/gif',
    intrinsicSize: { widthPt: 6, heightPt: 6 },
  }];
}

describe('paint resource registry', () => {
  it('preserves signed unbounded finite DrawingML source-rectangle percentages', () => {
    const authored = { l: -0.25, t: 1.25, r: 1.5, b: -0.75 };
    const srcRect = { ...authored };
    const registry = createPaintResourceRegistry([{
      kind: 'image',
      resourceKey: 'image:signed-crop',
      partPath: 'word/media/image.png',
      mimeType: 'image/png',
      intrinsicSize: { widthPt: 40, heightPt: 30 },
      srcRect,
    }]);
    srcRect.l = 0;

    const retained = registry.descriptors[0]!;
    expect(retained.kind).toBe('image');
    if (retained.kind !== 'image' && retained.kind !== 'picture-bullet') {
      throw new Error('expected an image paint resource');
    }
    const cloned = structuredClone(retained);

    expect(retained.srcRect).toEqual(authored);
    expect(cloned.srcRect).toEqual(authored);
    expect(stableFingerprint('paint-resource', retained))
      .toBe(stableFingerprint('paint-resource', cloned));
    expect(stableFingerprint('paint-resource', retained))
      .not.toBe(stableFingerprint('paint-resource', {
        ...retained,
        srcRect: { l: 0, t: 1, r: 1, b: 0 },
      }));
  });

  it('deep snapshots and freezes sorted structured-clone-safe descriptors', () => {
    const input = descriptors();
    const registry = createPaintResourceRegistry(input);

    (input[1] as unknown as { intrinsicSize: { widthPt: number } }).intrinsicSize.widthPt = 999;
    (input[2] as unknown as { model: { title: string } }).model.title = 'Mutated';

    expect(registry.keys).toEqual([
      'chart:body:0',
      'image:body:0:word%2Fmedia%2Fimage1.png',
      'image:body:1:word%2Fmedia%2Fbullet.gif',
      'math:body:0:inline',
    ]);
    expect(registry.resolve('image:body:0:word%2Fmedia%2Fimage1.png', 'image'))
      .toMatchObject({
        partPath: 'word/media/image1.png',
        svgImagePath: 'word/media/image1.svg',
        intrinsicSize: { widthPt: 40 },
        rotation: 15,
        flipH: true,
        flipV: true,
        alpha: 0.5,
      });
    expect(registry.resolve('chart:body:0', 'chart').model.title).toBe('Original');
    expect(Object.isFrozen(registry.keys)).toBe(true);
    expect(Object.isFrozen(registry.descriptors)).toBe(true);
    expect(Object.isFrozen(registry.resolve('chart:body:0', 'chart').model)).toBe(true);
    expect(structuredClone(registry.descriptors)).toEqual(registry.descriptors);
  });

  it('rejects duplicate keys before one descriptor can shadow another', () => {
    const duplicate = descriptors()[0]!;
    expect(() => createPaintResourceRegistry([duplicate, { ...duplicate }]))
      .toThrow(/Duplicate paint resource key.*math:body:0:inline/);
  });

  it('throws for missing keys and kind mismatches', () => {
    const registry = createPaintResourceRegistry(descriptors());

    expect(() => registry.resolve('missing', 'image'))
      .toThrow(/Unknown paint resource: missing/);
    expect(() => registry.resolve('chart:body:0', 'image'))
      .toThrow(/Paint resource kind mismatch.*chart:body:0.*expected image.*chart/);
  });

  it('rejects non-plain or non-cloneable descriptor payloads', () => {
    const invalidModel = chartModel('Invalid') as ChartModel & { callback?: () => void };
    invalidModel.callback = () => undefined;
    expect(() => createPaintResourceRegistry([{
      kind: 'chart', resourceKey: 'chart:invalid',
      intrinsicSize: { widthPt: 1, heightPt: 1 }, model: invalidModel,
    }])).toThrow(/structured-clone-safe plain data/i);

    const mapModel = chartModel('Map') as ChartModel & { lookup?: Map<string, string> };
    mapModel.lookup = new Map([['a', 'b']]);
    expect(() => createPaintResourceRegistry([{
      kind: 'chart', resourceKey: 'chart:map',
      intrinsicSize: { widthPt: 1, heightPt: 1 }, model: mapModel,
    }])).toThrow(/structured-clone-safe plain data/i);
  });

  it.each([
    [{ kind: 'math', resourceKey: '' }, /resourceKey.*non-empty/i],
    [{
      kind: 'image', resourceKey: 'image:empty-path', partPath: '', mimeType: 'image/png',
      intrinsicSize: { widthPt: 1, heightPt: 1 },
    }, /partPath.*non-empty/i],
    [{
      kind: 'picture-bullet', resourceKey: 'image:empty-mime', partPath: 'bullet.png', mimeType: '',
      intrinsicSize: { widthPt: 1, heightPt: 1 },
    }, /mimeType.*non-empty/i],
    [{
      kind: 'image', resourceKey: 'image:negative', partPath: 'image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: -1, heightPt: 1 },
    }, /widthPt.*finite and non-negative/i],
    [{
      kind: 'chart', resourceKey: 'chart:nan', intrinsicSize: { widthPt: 1, heightPt: Number.NaN },
      model: chartModel('NaN'),
    }, /heightPt.*finite and non-negative/i],
    [{
      kind: 'image', resourceKey: 'image:alpha', partPath: 'image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 1, heightPt: 1 }, alpha: 1.1,
    }, /alpha.*between 0 and 1/i],
    [{
      kind: 'image', resourceKey: 'image:crop-edge', partPath: 'image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 1, heightPt: 1 }, srcRect: { l: Number.NaN, t: 0, r: 0, b: 0 },
    }, /srcRect.l.*finite/i],
    [{
      kind: 'image', resourceKey: 'image:crop-edge', partPath: 'image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 1, heightPt: 1 },
      srcRect: { l: 0, t: 0, r: 0, b: Number.POSITIVE_INFINITY },
    }, /srcRect.b.*finite/i],
    [{
      kind: 'image', resourceKey: 'image:rotation', partPath: 'image.png', mimeType: 'image/png',
      intrinsicSize: { widthPt: 1, heightPt: 1 }, rotation: Number.POSITIVE_INFINITY,
    }, /rotation.*finite/i],
  ] as const)('rejects invalid descriptor contract %#', (descriptor, message) => {
    expect(() => createPaintResourceRegistry([
      descriptor as unknown as PaintResourceDescriptor,
    ])).toThrow(message);
  });
});
