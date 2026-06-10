// Public API of the bidirectional-text (UAX#9) module. Renderers import only
// from here.

export { REMOVED_LEVEL } from './types.js';
export type {
  BaseDirection,
  BidiClass,
  BidiLevels,
  StyledRun,
  VisualSegment,
  SegmentPart,
} from './types.js';

export type { BidiEngine } from './engine.js';
export {
  getDefaultBidiEngine,
  setBidiEngine,
  resetBidiEngine,
} from './engine.js';

export { toVisualSegments, resolveBaseDirection } from './segments.js';
