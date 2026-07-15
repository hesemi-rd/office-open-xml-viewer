import { describe, expect, it } from 'vitest';
import { paragraphAcquisitionInput, vmlTextPathAcquisitionInput } from './parser-model.js';
import type { DocParagraph, ShapeRun } from './types.js';

describe('private VML textPath parser facts', () => {
  it('projects resolved parser wire facts into an immutable acquisition input', () => {
    const parserShape = {
      textPath: {
        string: 'DRAFT',
        fontFamily: 'Calibri',
        bold: false,
        italic: false,
        textPathOk: true,
        on: true,
        fitShape: true,
        fitPath: false,
        trim: false,
        xScale: false,
        fontSizePt: 1,
      },
    } as unknown as ShapeRun;

    const input = vmlTextPathAcquisitionInput(parserShape);

    expect(input).toEqual({
      string: 'DRAFT',
      fontFamily: 'Calibri',
      bold: false,
      italic: false,
      textPathOk: true,
      on: true,
      fitShape: true,
      fitPath: false,
      trim: false,
      xScale: false,
      fontSizePt: 1,
    });
    expect(Object.isFrozen(input)).toBe(true);
  });

  it('distinguishes parser default-false controls from a manual public fallback', () => {
    const parserShape = {
      textPath: {
        string: 'NOTICE',
        textPathOk: false,
        on: false,
        fitShape: false,
        fitPath: false,
        trim: false,
        xScale: false,
      },
    } as unknown as ShapeRun;
    const publicFallback = {
      textPath: { string: 'NOTICE' },
    } as unknown as ShapeRun;

    const parserInput = vmlTextPathAcquisitionInput(parserShape);
    const publicInput = vmlTextPathAcquisitionInput(publicFallback);

    expect(parserInput).toMatchObject({
      textPathOk: false,
      on: false,
      fitShape: false,
      fitPath: false,
      trim: false,
      xScale: false,
    });
    for (const key of ['textPathOk', 'on', 'fitShape', 'fitPath', 'trim', 'xScale']) {
      expect(publicInput).not.toHaveProperty(key);
    }
    expect(publicInput).not.toHaveProperty('fontFamily');
    expect(publicInput).not.toHaveProperty('fontSizePt');
    expect(JSON.parse(JSON.stringify(parserInput))).toEqual(parserInput);
    expect(JSON.parse(JSON.stringify(publicInput))).toEqual(publicInput);
  });

  it('attaches parser-only controls to the immutable paragraph shape snapshot', () => {
    const shape = {
      type: 'shape',
      textPath: {
        string: 'ARCHIVE',
        textPathOk: true,
        on: true,
        fitShape: false,
        fitPath: false,
        trim: true,
        xScale: false,
        fontSizePt: 14,
      },
    } as unknown as ShapeRun & { type: 'shape' };
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], runs: [shape],
    } as DocParagraph;

    const snapshot = paragraphAcquisitionInput(paragraph, {
      story: 'body', storyInstance: 'body', path: [0],
    });
    const acquiredShape = snapshot.runs[0];

    expect(acquiredShape).toMatchObject({
      type: 'shape',
      vmlTextPathInput: {
        string: 'ARCHIVE', textPathOk: true, on: true, fitShape: false,
        fitPath: false, trim: true, xScale: false, fontSizePt: 14,
      },
    });
    expect(Object.isFrozen(acquiredShape)).toBe(true);
    expect(Object.isFrozen(
      acquiredShape.type === 'shape' ? acquiredShape.vmlTextPathInput : undefined,
    )).toBe(true);
  });
});
