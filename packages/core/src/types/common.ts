// Format-agnostic OOXML types shared across pptx/docx/xlsx packages.
// All positions and sizes are in EMUs (English Metric Units).

export type PathCmd =
  | { cmd: 'moveTo';     x: number; y: number }
  | { cmd: 'lineTo';     x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'arcTo';      wr: number; hr: number; stAng: number; swAng: number }
  | { cmd: 'close' };

export type Fill = SolidFill | NoFill | GradientFill | PatternFill;

export interface SolidFill {
  fillType: 'solid';
  color: string; // hex 6-char or 8-char (RRGGBBAA with alpha)
}

export interface NoFill {
  fillType: 'none';
}

export interface GradientStop {
  position: number; // 0.0–1.0
  color: string;    // hex 6 or 8 chars
}

export interface GradientFill {
  fillType: 'gradient';
  stops: GradientStop[];
  /** degrees: 0 = left→right, 90 = top→bottom */
  angle: number;
  /** 'linear' | 'radial' */
  gradType: string;
}

/**
 * Preset pattern fill — ECMA-376 §20.1.8.40 (CT_PatternFillProperties)
 * with `preset` drawn from §20.1.10.59 (ST_PresetPatternVal).
 */
export interface PatternFill {
  fillType: 'pattern';
  /** Foreground hex colour — used for the "1" pixels of the preset bitmap. */
  fg: string;
  /** Background hex colour — used for the "0" pixels. */
  bg: string;
  /** Preset name, e.g. "pct25", "horz", "diagCross", "lgGrid". */
  preset: string;
}

export interface Shadow {
  color: string;  // hex 6 chars
  alpha: number;  // 0.0–1.0
  blur: number;   // EMU
  dist: number;   // EMU
  /** degrees clockwise from East */
  dir: number;
}

/** ECMA-376 §20.1.8.17 (CT_GlowEffect) — coloured halo with blur radius. */
export interface Glow {
  color: string;  // hex 6 chars
  alpha: number;  // 0.0–1.0
  /** Blur radius in EMU. */
  radius: number;
}

/** ECMA-376 §20.1.8.31 (CT_SoftEdgesEffect) — feather radius in EMU. */
export interface SoftEdge {
  radius: number;
}

/** ECMA-376 §20.1.8.27 (CT_ReflectionEffect) — mirrored copy below the
 *  shape with a linear alpha gradient. Carries the spec attributes whose
 *  defaults the renderer needs to interpret correctly. */
export interface Reflection {
  blur: number;      // EMU
  dist: number;      // EMU
  /** Direction in degrees, clockwise from East. */
  dir: number;
  /** Start alpha (0–1). Default 1.0. */
  stA: number;
  /** Start position along the gradient (0–1). Default 0. */
  stPos: number;
  /** End alpha. Default 0. */
  endA: number;
  /** End position. Default 1.0. */
  endPos: number;
  /** Horizontal scale (1.0 = same width). */
  sx: number;
  /** Vertical scale (-1.0 = full mirror). */
  sy: number;
}

export interface ArrowEnd {
  /** OOXML type: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow" */
  type: string;
  /** Width multiplier: "sm" | "med" | "lg" */
  w: string;
  /** Length multiplier: "sm" | "med" | "lg" */
  len: string;
}

export interface Stroke {
  color: string;
  /** Width in EMU */
  width: number;
  /** OOXML prstDash value: "dash", "dot", "dashDot", "lgDash", "lgDashDot", etc. */
  dashStyle?: string;
  /** Arrow head at the start of the line */
  headEnd?: ArrowEnd;
  /** Arrow head at the end of the line */
  tailEnd?: ArrowEnd;
  /**
   * ECMA-376 §20.1.8.42 ST_CompoundLine. "sng" (default) | "dbl" |
   * "thinThick" | "thickThin" | "tri". Absent means single line.
   */
  cmpd?: string;
}

export interface TextBody {
  /** Vertical anchor: "t" | "ctr" | "b" */
  verticalAnchor: string;
  paragraphs: Paragraph[];
  /** Default pt size from lstStyle (overrides renderer default when present) */
  defaultFontSize: number | null;
  /** Inherited bold from layout/master defRPr (null = not set, use false as final default) */
  defaultBold: boolean | null;
  /** Inherited italic from layout/master defRPr (null = not set, use false as final default) */
  defaultItalic: boolean | null;
  /** Text insets in EMU (defaults: lIns=rIns=91440, tIns=bIns=45720) */
  lIns: number;
  rIns: number;
  tIns: number;
  bIns: number;
  /** "square" = wrap, "none" = no wrap */
  wrap: string;
  /** Text direction: "horz" | "vert" | "vert270" | "eaVert" etc. */
  vert: string;
  /** Auto-fit: "sp" = shape grows to fit text, "norm" = font shrinks, "none" = no fit */
  autoFit: string;
  /**
   * `<a:normAutofit fontScale>` (ECMA-376 §21.1.2.1.3) — PowerPoint's stored,
   * pre-computed font-shrink ratio for `autoFit === "norm"`, as a fraction
   * (e.g. 0.625 for `fontScale="62500"`). Null/absent when PowerPoint stored no
   * scale; the renderer then re-derives one. Applying the stored value matches
   * PowerPoint exactly instead of guessing from our own text metrics.
   */
  fontScale?: number | null;
  /** `<a:normAutofit lnSpcReduction>` — stored line-spacing reduction fraction
   *  (e.g. 0.20 for `lnSpcReduction="20000"`). Null/absent when not stored. */
  lnSpcReduction?: number | null;
  /**
   * `<a:bodyPr numCol>` (ECMA-376 §20.1.10.34) — number of text columns inside
   * the shape. Defaults to 1; values > 1 cause the renderer to flow paragraphs
   * across N columns left-to-right, top-to-bottom.
   */
  numCol?: number;
  /** `<a:bodyPr spcCol>` — gap between columns in EMU. Default 0. */
  spcCol?: number;
}

export type SpaceLine =
  | { type: 'pct'; val: number }   // val: e.g. 100000 = 100%, 150000 = 150%
  | { type: 'pts'; val: number };  // val in points

export type Bullet =
  | { type: 'none' }
  | { type: 'inherit' }
  | { type: 'char'; char: string; color: string | null; sizePct: number | null; fontFamily: string | null }
  | { type: 'autoNum'; numType: string; startAt: number | null };

export interface TabStop {
  /** Position in EMU from the left edge of the text area (after lIns) */
  pos: number;
  /** Alignment: "l" | "r" | "ctr" | "dec" */
  algn: string;
}

export interface Paragraph {
  /** Alignment: "l" | "ctr" | "r" | "just" */
  alignment: string;
  /** Left margin in EMU */
  marL: number;
  /** Right margin in EMU */
  marR: number;
  /** First-line indent in EMU (negative = hanging indent) */
  indent: number;
  spaceBefore: number | null;
  spaceAfter: number | null;
  spaceLine: SpaceLine | null;
  /** List nesting level (0–8) */
  lvl: number;
  bullet: Bullet;
  defFontSize: number | null;
  defColor: string | null;
  defBold: boolean | null;
  defItalic: boolean | null;
  defFontFamily: string | null;
  /** Tab stops from pPr > tabLst */
  tabStops: TabStop[];
  /**
   * `<a:pPr rtl="1">` — right-to-left paragraph (ECMA-376 §21.1.2.2.7).
   * When true and no explicit `algn`, the parser-side default flips from
   * "l" to "r"; renderers can also use this flag to flow runs RTL.
   */
  rtl?: boolean;
  runs: TextRun[];
}

export type TextRun = TextRunData | LineBreak;

export interface TextRunData {
  type: 'text';
  text: string;
  /** null = not set, inherit from paragraph/body defaults */
  bold: boolean | null;
  /** null = not set, inherit from paragraph/body defaults */
  italic: boolean | null;
  underline: boolean;
  /**
   * Specific underline style when not the default single line. Values come
   * from ECMA-376 §21.1.2.3.16 (ST_TextUnderlineType): "dbl", "heavy",
   * "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong",
   * "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash",
   * "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDbl". Absent means either
   * no underline (when `underline` is false) or the default single line.
   */
  underlineStyle?: string;
  /**
   * Underline-only colour from rPr > uFill (ECMA-376 §21.1.2.3.20). Absent
   * means the underline follows the text colour (uFillTx default).
   */
  underlineColor?: string;
  /** True when rPr strike is sngStrike or dblStrike. */
  strikethrough: boolean;
  /**
   * True only when rPr strike = "dblStrike". Lets the renderer draw two parallel
   * lines instead of one. ECMA-376 §21.1.2.3.10 (ST_TextStrikeType).
   */
  strikeDouble?: boolean;
  /** Font size in points */
  fontSize: number | null;
  color: string | null;
  fontFamily: string | null;
  /**
   * East Asian font family from rPr > a:ea (ECMA-376 §21.1.2.3.7),
   * resolved through the theme. Renderer uses this for CJK glyphs when
   * present; absent means CJK falls back to fontFamily.
   */
  fontFamilyEa?: string;
  /** Baseline shift in thousandths of a point. Positive = superscript, negative = subscript. */
  baseline?: number;
  /**
   * Capitalisation transform — ECMA-376 §21.1.2.3.13 (ST_TextCapsType).
   * 'all' renders text in upper case; 'small' uses small caps (rendered as
   * upper case at ~80% size when no smcp font feature is available).
   * 'none' or omitted leaves the text unchanged.
   */
  caps?: 'none' | 'small' | 'all';
  /**
   * Inter-character spacing in 100ths of a point — ECMA-376 §21.1.2.3.5
   * (rPr @spc). Positive values add space, negative values tighten.
   */
  letterSpacing?: number;
  /** Set for OOXML field runs (e.g. "slidenum"). When set, renderer replaces text with field value. */
  fieldType?: string;
  /**
   * Hyperlink target URL resolved from rPr > a:hlinkClick @r:id via the slide's _rels.
   * Undefined for runs without a hyperlink. ECMA-376 §21.1.2.3.5 (CT_Hyperlink).
   */
  hyperlink?: string;
  /**
   * Run-level drop shadow on glyphs (`<a:rPr><a:effectLst><a:outerShdw>`),
   * ECMA-376 §20.1.8.45. Independent of the shape-level shadow on `spPr`.
   * Absent means no run-level shadow.
   */
  shadow?: Shadow;
  /**
   * Run-level glyph outline (`<a:rPr><a:ln w="..">`), ECMA-376 §20.1.2.2.24
   * (CT_TextOutlineEffect). Renderer strokes each glyph with the given
   * width / colour in addition to the normal fill. Absent means glyphs are
   * fill-only.
   */
  outline?: TextOutline;
}

/** Run-level glyph outline. Width is in OOXML EMU (12700 EMU = 1 pt). */
export interface TextOutline {
  width: number;
  /** Hex without '#'. Absent = inherit from text fill colour. */
  color?: string;
}

export interface LineBreak {
  type: 'break';
}

export interface RenderOptions {
  width?: number;
  defaultTextColor?: string | null;
  dpr?: number;
  majorFont?: string | null;
  minorFont?: string | null;
  /** Theme hyperlink colour (hex 6 chars). Used to colour hyperlink runs without an explicit colour. */
  hlinkColor?: string | null;
  /**
   * Lazily resolve an archive-internal asset (by zip path) to a Blob. The
   * renderer uses this to fetch posters and other large embedded assets on
   * demand, keeping the parse output free of inlined base64.
   */
  fetchMedia?: (path: string) => Promise<Blob>;
  /**
   * When true, renderMedia draws only the poster frame — play/pause badges
   * and progress bars are left to the caller. Set by the pptx presentSlide
   * API so its interactive handle can own all control chrome without
   * the static renderer drawing a duplicate play badge.
   */
  skipMediaControls?: boolean;
}
