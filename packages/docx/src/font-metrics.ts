// Font design line-metrics moved to the shared core layer so docx, pptx and
// xlsx size line boxes from the SAME OS/2 win / hhea table (a substituted
// Meiryo / Sakkal Majalla / Times / Arial face must measure to the intended
// font's design line height in all three formats). This thin re-export keeps
// docx's existing import sites (`renderer.ts`, tests) unchanged.
//
// See `@silurus/ooxml-core` → `text/line-metrics.ts` for the table, provenance,
// and the win-vs-hhea rationale.
export {
  fontWinLineHeightRatio,
  intendedSingleLinePx,
  correctLineMetrics,
} from '@silurus/ooxml-core';
