export type {
  Change,
  ChangeOp,
  ChangeLocation,
  BBox,
  DiffResult,
  Format,
} from './types.ts';
export { changesAtSlide, changesAtPage, changesAtSheet } from './types.ts';

export { diffPptx, bboxesForSlide } from './pptx.ts';
export { diffDocx } from './docx.ts';
export { diffXlsx, type XlsxDiffInput } from './xlsx.ts';

export { alignSequences, type SequenceAlignment } from './util/sequence.ts';
export { deepEqual } from './util/equal.ts';
