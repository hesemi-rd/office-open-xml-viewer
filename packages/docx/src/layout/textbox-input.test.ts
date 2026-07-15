import { describe, expect, it } from 'vitest';
import { normalizeTextBoxInput } from './textbox-input.js';
import { textBoxAcquisitionInput } from '../parser-model.js';
import type { ShapeRun } from '../types.js';

describe('normalizeTextBoxInput', () => {
  it('converts public textBlocks to parser-independent paragraph inputs in source order', () => {
    const shape = {
      type: 'shape',
      textBlocks: [
        { text: 'first', fontSizePt: 11, color: '112233', alignment: 'right',
          indentLeft: 4, indentRight: 5, indentFirst: -2,
          lineSpacingVal: 14, lineSpacingRule: 'exact',
          tabStops: [{ pos: 24, alignment: 'right', leader: 'dot' }],
          bidi: true, contextualSpacing: true, styleId: 'Quote' },
        { text: 'second', fontSizePt: 12, color: '445566' },
      ],
    } as unknown as ShapeRun;

    const inputs = normalizeTextBoxInput(shape, {
      story: 'textbox', storyInstance: 'shape:4', path: [4],
    });

    expect(inputs).toHaveLength(2);
    expect(inputs.map((input) => input.source.path)).toEqual([[4, 0], [4, 1]]);
    expect(inputs.map((input) => input.runs[0])).toEqual([
      expect.objectContaining({ text: 'first', color: '#112233' }),
      expect.objectContaining({ text: 'second', color: '#445566' }),
    ]);
    expect(inputs.every((input) => !('shape' in input) && !('lines' in input))).toBe(true);
    expect(inputs[0]).toMatchObject({
      alignment: 'right', indentLeftPt: 4, indentRightPt: 5, indentFirstPt: -2,
      lineSpacing: { value: 14, rule: 'exact', explicit: true },
      tabStops: [{ pos: 24, alignment: 'right', leader: 'dot' }],
      bidi: true, contextualSpacing: true, styleId: 'Quote',
    });
  });

  it('returns deeply frozen plain data without aliasing or freezing caller-owned nested values', () => {
    const ruby = { text: 'かん', fontSizePt: 6, hpsRaisePt: 2 };
    const fontFacts = {
      fontSize: 13,
      fontSlots: {
        direct: { ascii: 'Caller Sans', highAnsi: 'Caller Sans', eastAsia: 'Caller Gothic', complexScript: null },
        theme: { ascii: null, highAnsi: null, eastAsia: null, complexScript: null },
        themePresent: { ascii: false, highAnsi: false, eastAsia: false, complexScript: false },
      },
    };
    const numbering = {
      numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 18, tab: 18, suff: 'tab',
      fontFacts,
    };
    const tabStop = { pos: 24, alignment: 'right' as const, leader: 'dot' as const };
    const runs = [{ text: '漢', fontSizePt: 12, ruby }];
    const textBlocks = [{
      text: '漢', fontSizePt: 12, alignment: 'left', spaceBefore: 3, spaceAfter: 4,
      runs, numbering, tabStops: [tabStop],
      imagePath: 'word/media/image.png', mimeType: 'image/png',
      imageWidthPt: 20, imageHeightPt: 10,
    }];
    const shape = { type: 'shape', textBlocks } as unknown as ShapeRun;
    const source = { story: 'textbox' as const, storyInstance: 'shape:7', path: [7] };

    const inputs = normalizeTextBoxInput(shape, source);

    expect(Object.isFrozen(textBlocks)).toBe(false);
    expect(Object.isFrozen(runs)).toBe(false);
    expect(Object.isFrozen(ruby)).toBe(false);
    expect(Object.isFrozen(numbering)).toBe(false);
    expect(Object.isFrozen(fontFacts)).toBe(false);
    expect(Object.isFrozen(source)).toBe(false);
    expect(Object.isFrozen(source.path)).toBe(false);
    expect(Object.isFrozen(tabStop)).toBe(false);

    ruby.text = 'へん';
    numbering.text = '9.';
    fontFacts.fontSlots.direct.ascii = 'Mutated Face';
    source.path[0] = 99;
    tabStop.pos = 99;
    textBlocks[0]!.spaceAfter = 99;

    expect(inputs[0]).toMatchObject({
      source: { path: [7, 0] },
      spacing: { beforePt: 3, afterPt: 4 },
      runs: [{ ruby: { text: 'かん' } }],
      numbering: { text: '1.', fontFacts: { fontSlots: { direct: { ascii: 'Caller Sans' } } } },
      tabStops: [{ pos: 24 }],
      image: { imagePath: 'word/media/image.png', widthPt: 20, heightPt: 10 },
    });
    expect(structuredClone(inputs)).toEqual(inputs);
    for (const value of [
      inputs, inputs[0], inputs[0]!.source, inputs[0]!.source.path,
      inputs[0]!.spacing, inputs[0]!.runs, inputs[0]!.runs[0], inputs[0]!.runs[0]!.ruby,
      inputs[0]!.numbering, (inputs[0]!.numbering as typeof numbering).fontFacts,
      inputs[0]!.tabStops, inputs[0]!.tabStops[0], inputs[0]!.numberingMarkerShapeInput,
      inputs[0]!.numberingMarkerShapeInput?.fonts, inputs[0]!.image,
    ]) expect(Object.isFrozen(value)).toBe(true);
  });

  it('replaces the compatibility marker from parser facts without freezing parser-owned data', () => {
    const ruby = { text: 'ふり', fontSizePt: 5 };
    const direct = {
      ascii: 'Parser Sans', highAnsi: 'Parser Sans', eastAsia: 'Parser Gothic', complexScript: null,
    };
    const fontFacts = {
      fontSize: 15, bold: true,
      fontSlots: {
        direct,
        theme: { ascii: null, highAnsi: null, eastAsia: null, complexScript: null },
        themePresent: { ascii: false, highAnsi: false, eastAsia: false, complexScript: false },
      },
    };
    const numbering = {
      numId: 2, level: 0, format: 'bullet', text: '•', indentLeft: 18, tab: 18, suff: 'tab',
      fontFacts,
    };
    const shape = {
      type: 'shape',
      textBlocks: [{
        text: '漢', fontSizePt: 10, alignment: 'left', numbering,
        runs: [{ text: '漢', fontSizePt: 10, ruby }],
      }],
    } as unknown as ShapeRun;
    const source = { story: 'textbox' as const, storyInstance: 'parser-shape', path: [2] };

    const inputs = textBoxAcquisitionInput(shape, source);

    expect(inputs[0]?.numberingMarkerShapeInput).toMatchObject({
      fontSizePt: 15,
      fonts: { ascii: 'Parser Sans', eastAsia: 'Parser Gothic' },
      weight: 700,
    });
    expect(Object.isFrozen(ruby)).toBe(false);
    expect(Object.isFrozen(numbering)).toBe(false);
    expect(Object.isFrozen(fontFacts)).toBe(false);
    expect(Object.isFrozen(direct)).toBe(false);
    ruby.text = '変更';
    numbering.text = 'x';
    direct.ascii = 'Changed';
    source.path[0] = 8;
    expect(inputs[0]).toMatchObject({
      source: { path: [2, 0] },
      runs: [{ ruby: { text: 'ふり' } }],
      numbering: { text: '•' },
      numberingMarkerShapeInput: { fonts: { ascii: 'Parser Sans' } },
    });
    expect(structuredClone(inputs)).toEqual(inputs);
    expect(Object.isFrozen(inputs[0]?.numberingMarkerShapeInput?.fonts)).toBe(true);
  });
});
