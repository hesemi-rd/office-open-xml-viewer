import { describe, expect, it } from 'vitest';
import {
  anchorAcquisitionInput,
  paragraphAcquisitionInput,
} from './parser-model.js';
import type { DocParagraph, ImageRun } from './types.js';

const privateWire = {
  occurrenceId: 'wp-anchor-120',
  simplePosition: {
    enabled: false, status: 'valid',
    xPt: 1, xStatus: 'valid', yPt: 2, yStatus: 'valid',
  },
  horizontal: { relativeFrom: 'page', relativeFromStatus: 'valid', choice: { kind: 'percent', fraction: 0.25 } },
  vertical: { relativeFrom: 'paragraph', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 3 } },
  extent: { widthPt: 20, heightPt: 10, widthStatus: 'valid', heightStatus: 'valid' },
  parentEffectExtent: { topPt: 1, topStatus: 'valid', rightPt: 2, rightStatus: 'valid', bottomPt: 3, bottomStatus: 'valid', leftPt: 4, leftStatus: 'valid' },
  anchorDistances: { topPt: 5, topStatus: 'valid', rightPt: 6, rightStatus: 'valid', bottomPt: 7, bottomStatus: 'valid', leftPt: 8, leftStatus: 'valid' },
  relativeSize: {
    horizontal: { relativeFrom: 'page', relativeFromStatus: 'valid', fraction: 0, fractionStatus: 'valid' },
    vertical: { relativeFrom: 'margin', relativeFromStatus: 'valid', fraction: 0.5, fractionStatus: 'valid' },
  },
  wrap: {
    kind: 'tight', authoredKinds: ['wrapTight'], side: 'largest',
    distances: { topPt: null, topStatus: 'missing', rightPt: 0.6, rightStatus: 'valid', bottomPt: null, bottomStatus: 'missing', leftPt: 0.5, leftStatus: 'valid' },
    effectExtent: null,
    polygon: {
      edited: true,
      coordinateSpace: { width: 21600, height: 21600 },
      points: [
        { x: 0, y: 0, rawX: '0', rawY: '0' },
        { x: 21600, y: 0, rawX: '21600', rawY: '0' },
        { x: 24000, y: 21600, rawX: '24000', rawY: '21600' },
      ],
      invalidPointCount: 0,
    },
  },
  behavior: {
    behindDoc: true, behindDocStatus: 'valid',
    relativeHeight: 42, relativeHeightStatus: 'valid',
    locked: false, lockedStatus: 'valid',
    allowOverlap: null, allowOverlapStatus: 'missing',
    layoutInCell: true, layoutInCellStatus: 'valid',
  },
  group: null,
} as const;

describe('private anchor acquisition projection', () => {
  it('projects parser-only facts as deeply frozen structured-clone-safe data', () => {
    const run = {
      type: 'image',
      __anchorAcquisition: privateWire,
    } as unknown as ImageRun & { type: 'image' };

    const input = anchorAcquisitionInput(run);

    expect(input).toEqual(privateWire);
    expect(input).not.toBe(privateWire);
    expect(structuredClone(input)).toEqual(input);
    expect(Object.isFrozen(input)).toBe(true);
    expect(Object.isFrozen(input?.wrap.polygon?.points)).toBe(true);
    expect(input?.relativeSize.horizontal?.fraction).toBe(0);
    expect(input?.behavior).toMatchObject({
      allowOverlapStatus: 'missing',
      allowOverlap: null,
      lockedStatus: 'valid',
      locked: false,
    });
  });

  it('distinguishes a public hand-built run from parser-produced private facts', () => {
    expect(anchorAcquisitionInput({ type: 'image' } as unknown as ImageRun)).toBeUndefined();
  });

  it('snapshots the parser-resolved grouped child frame without source aliasing', () => {
    const resolvedChildFrame = {
      offsetXPt: 1, offsetYPt: 2, widthPt: 3, heightPt: 4,
      rotationDeg: 30, flipH: true, flipV: false,
    };
    const wire = {
      ...privateWire,
      group: {
        childSourceId: 'group-child-1', sourceIndex: 0, sourceCount: 1,
        transformChain: [], childTransform: null, resolvedChildFrame,
      },
    };
    const input = anchorAcquisitionInput({
      type: 'image', __anchorAcquisition: wire,
    } as unknown as ImageRun);

    resolvedChildFrame.offsetXPt = 99;

    expect(input?.group?.resolvedChildFrame.offsetXPt).toBe(1);
    expect(Object.isFrozen(input?.group?.resolvedChildFrame)).toBe(true);
    expect(structuredClone(input)).toEqual(input);
  });

  it('attaches the same immutable projection to paragraph image snapshots', () => {
    const host = {
      type: 'anchorHost', fontSize: 11,
      __anchorOccurrenceId: privateWire.occurrenceId,
    };
    const run = {
      type: 'image',
      __anchorAcquisition: privateWire,
    } as unknown as ImageRun & { type: 'image' };
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], runs: [host, run],
    } as DocParagraph;

    const snapshot = paragraphAcquisitionInput(paragraph, {
      story: 'body', storyInstance: 'body', path: [0],
    });
    const acquiredHost = snapshot.runs[0] as unknown as { anchorOccurrenceId?: string };
    const acquired = snapshot.runs[1] as unknown as {
      anchorAcquisitionInput?: unknown;
    };

    const scopedOccurrenceId = 'anchor:body:body:0:wp-anchor-120';
    expect(acquiredHost.anchorOccurrenceId).toBe(scopedOccurrenceId);
    expect(acquired.anchorAcquisitionInput).toEqual({
      ...privateWire,
      occurrenceId: scopedOccurrenceId,
    });
    expect(Object.isFrozen(acquired.anchorAcquisitionInput)).toBe(true);
  });
});
