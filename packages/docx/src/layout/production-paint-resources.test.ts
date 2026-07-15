import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import type { DocParagraph, DocxDocumentModel } from '../types.js';
import { imageResourceKey, mathResourceKey } from './source-key.js';
import {
  chartPaintResourceKey,
  createDocumentPaintResourceRegistry,
} from './production-paint-resources.js';

function chartModel(): ChartModel {
  return {
    chartType: 'line', title: null, categories: [], series: [], varyColors: false,
    showDataLabels: false, valMin: null, valMax: null, catAxisTitle: null,
    valAxisTitle: null, catAxisHidden: false, valAxisHidden: false,
    catAxisLineHidden: false, valAxisLineHidden: false, plotAreaBg: null,
    chartBg: null, showLegend: false, legendPos: null, catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'cross', catAxisMajorTickMark: 'cross',
    titleFontSizeHpt: null, titleFontColor: null, titleFontFace: null,
    catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null, subtotalIndices: [],
  } as ChartModel;
}

function paragraph(runs: DocParagraph['runs']): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
    numbering: {
      numId: 1, level: 0, format: 'bullet', text: '', indentLeft: 0, tab: 18,
      suff: 'tab', picBulletImagePath: 'word/media/bullet.gif',
      picBulletMimeType: 'image/gif', picBulletWidthPt: 6, picBulletHeightPt: 7,
    },
    runs,
  } as DocParagraph;
}

function documentModel(): DocxDocumentModel {
  return {
    section: {} as DocxDocumentModel['section'],
    headers: {
      default: {
        body: [paragraph([{
          type: 'image', imagePath: 'word/media/header.png', mimeType: 'image/png',
          widthPt: 12, heightPt: 8,
        }])],
      },
    },
    footers: {},
    body: [paragraph([{
      type: 'image', imagePath: 'word/media/image.png', svgImagePath: 'word/media/image.svg',
      mimeType: 'image/png', widthPt: 40, heightPt: 30,
      srcRect: { l: .1, t: .2, r: .1, b: 0 }, rotation: 15, flipH: true,
      flipV: true, alpha: .5, colorReplaceFrom: 'FFFFFF',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    }, {
      type: 'chart', chart: chartModel(), widthPt: 120, heightPt: 80, anchor: false,
    }, {
      type: 'math', nodes: [], display: false, fontSize: 10,
    }, {
      type: 'shape', widthPt: 50, heightPt: 30, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0, subpaths: [],
      fill: null, stroke: null, textBlocks: [{
        text: '', fontSizePt: 10, alignment: 'left', imagePath: 'word/media/textbox.png',
        mimeType: 'image/png', imageWidthPt: 9, imageHeightPt: 5,
      }],
    }])],
  } as DocxDocumentModel;
}

describe('production paint resources', () => {
  it('builds clone-safe descriptors with the same structural keys as retained layout', () => {
    const registry = createDocumentPaintResourceRegistry(documentModel());
    const body = { story: 'body' as const, storyInstance: 'body', path: [0] };
    const imageKey = imageResourceKey({ ...body, path: [0, 0] }, 'word/media/image.png');
    const chartKey = chartPaintResourceKey({ ...body, path: [0, 1] });
    const mathKey = mathResourceKey({ ...body, path: [0, 2] }, 'inline');
    const bulletKey = imageResourceKey(body, 'word/media/bullet.gif');
    const textBoxKey = imageResourceKey({
      story: 'textbox', storyInstance: 'body:body:0.3', path: [0, 0],
    }, 'word/media/textbox.png');
    const headerImageKey = imageResourceKey({
      story: 'header', storyInstance: 'default', path: [0, 0],
    }, 'word/media/header.png');
    const headerBulletKey = imageResourceKey({
      story: 'header', storyInstance: 'default', path: [0],
    }, 'word/media/bullet.gif');

    expect(registry.keys).toEqual([
      bulletKey, headerBulletKey, headerImageKey, imageKey, textBoxKey, chartKey, mathKey,
    ].sort((left, right) => left.localeCompare(right)));
    expect(registry.resolve(imageKey, 'image')).toMatchObject({
      partPath: 'word/media/image.png', svgImagePath: 'word/media/image.svg',
      intrinsicSize: { widthPt: 40, heightPt: 30 },
      srcRect: { l: .1, t: .2, r: .1, b: 0 }, rotation: 15,
      flipH: true, flipV: true, alpha: .5, colorReplaceFrom: 'FFFFFF',
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    });
    expect(registry.resolve(chartKey, 'chart').model.chartType).toBe('line');
    expect(registry.resolve(bulletKey, 'picture-bullet')).toMatchObject({
      partPath: 'word/media/bullet.gif', intrinsicSize: { widthPt: 6, heightPt: 7 },
    });
    expect(structuredClone(registry.descriptors)).toEqual(registry.descriptors);
  });
});
