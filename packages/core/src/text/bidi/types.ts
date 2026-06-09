// Shared types for the bidirectional-text (UAX#9) module.
//
// This module turns a line's logical sequence of styled runs into a
// visual-ordered list of draw segments. Ordering / mirroring / level
// resolution follow the Unicode Bidirectional Algorithm (UAX#9); intra-segment
// glyph shaping and reversal are left to the Canvas text stack (each segment is
// drawn with one fillText and ctx.direction matching its resolved direction).

/** Paragraph base direction. `auto` = resolve by UAX#9 first-strong (P2-P3). */
export type BaseDirection = 'ltr' | 'rtl' | 'auto';

/**
 * UAX#9 bidirectional character classes (Bidi_Class property values).
 * Strong: L, R, AL. Weak: EN, ES, ET, AN, CS, NSM, BN. Neutral: B, S, WS, ON.
 * Explicit formatting: LRE, LRO, RLE, RLO, PDF, LRI, RLI, FSI, PDI.
 */
export type BidiClass =
  | 'L' | 'R' | 'AL'
  | 'EN' | 'ES' | 'ET' | 'AN' | 'CS' | 'NSM' | 'BN'
  | 'B' | 'S' | 'WS' | 'ON'
  | 'LRE' | 'LRO' | 'RLE' | 'RLO' | 'PDF'
  | 'LRI' | 'RLI' | 'FSI' | 'PDI';

/** Per-code-unit embedding levels produced by {@link BidiEngine.computeLevels}. */
export type BidiLevels = Uint8Array;

/**
 * One styled run as the renderer sees it. `meta` is opaque payload the renderer
 * gets back on each produced segment (color, decoration, vertical alignment, a
 * back-reference to its own run object, etc.).
 *
 * The four typed fields are the SHAPE-AFFECTING identity: a change in any of
 * them forces a segment split (Arabic joining cannot cross a font/weight/
 * italic/size boundary). Non-shape-affecting style (color, underline,
 * track-change) must live in `meta` so it does NOT split a word.
 */
export interface StyledRun {
  text: string;
  fontFamily: string;
  bold: boolean;
  italic: boolean;
  fontSizePx: number;
  meta?: unknown;
}

/**
 * One visual draw segment: a maximal slice that is direction-uniform AND
 * shape-style-uniform. Segments are returned in VISUAL (left-to-right) order;
 * the renderer assigns x left-to-right by measured width.
 *
 * `text` is the slice in LOGICAL order — pass it to a single `fillText` with
 * `ctx.direction = isRTL ? 'rtl' : 'ltr'` so Canvas applies the correct
 * intra-segment reversal, shaping and bracket mirroring.
 */
export interface VisualSegment {
  text: string;
  isRTL: boolean;
  level: number;
  run: StyledRun;
}
