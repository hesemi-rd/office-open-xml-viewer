import type {
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
  MediaElement,
  TableElement,
  Fill,
  Stroke,
  TextBody,
  Paragraph,
  TextRun,
  PathCmd,
  Shadow,
  Glow,
  RenderOptions,
  TabStop,
} from './types';
import {
  renderChart,
  buildCustomPath as buildCustomPathCore,
  hexToRgba as hexToRgbaCore,
  resolveFill as resolveFillCore,
  applyStroke as applyStrokeCore,
  buildShapePath,
  drawStar,
  drawPolygon,
  ooxmlArcTo,
  EMU_PER_PT as PT_TO_EMU,
  mathToMathML,
  recolorSvg,
  applyInnerShadow,
  applySoftEdge,
  applyReflection,
} from '@silurus/ooxml-core';
import type { MathNode, MathRenderer } from '@silurus/ooxml-core';
import { drawPlayBadge } from './media-chrome';
import { renderPresetShape, hasPreset, getConnectorAnchors } from './preset-shape';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  type LineVisualOrder,
} from './bidi-line';

/** Theme font context threaded through the render call chain. */
export interface RenderContext {
  themeMajorFont: string | null;
  themeMinorFont: string | null;
  /** Theme hyperlink colour as a 6-char hex (no leading #), or null. */
  themeHlinkColor?: string | null;
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface PptxTextRunInfo {
  text: string;
  /** X position in CSS px, relative to the shape's top-left corner. */
  inShapeX: number;
  /** Y position (top of line box) in CSS px, relative to the shape's top-left corner. */
  inShapeY: number;
  /** Measured text width in CSS px. */
  w: number;
  /** Line height in CSS px. */
  h: number;
  /** Font size in CSS px. */
  fontSize: number;
  /** CSS `font` shorthand used for canvas drawing (e.g. `"bold 16px Arial"`). */
  font: string;
  /** Shape's left edge in canvas CSS px. */
  shapeX: number;
  /** Shape's top edge in canvas CSS px. */
  shapeY: number;
  /** Shape's width in canvas CSS px. */
  shapeW: number;
  /** Shape's height in canvas CSS px. */
  shapeH: number;
  /** Shape rotation in degrees (clockwise). */
  rotation: number;
  /**
   * Additional rotation from a vertical text body (`vert="vert"` → 90,
   * `vert="vert270"` → -90). The CSS overlay must add this to `rotation`.
   */
  textBodyRotation?: number;
}

export type TextRunCallback = (run: PptxTextRunInfo) => void;

/**
 * Convert EMU to canvas pixels.
 * scale = canvasWidthPx / slideWidthEMU  (so that slideWidth EMU == canvasWidth px)
 */
function emuToPx(emu: number, scale: number): number {
  return emu * scale;
}

const hexToRgba = hexToRgbaCore;

/** Simple fill resolver that returns a CSS color string.
 *  For gradient/pattern fills, returns a flat colour (used by table cells etc.,
 *  where we don't have ctx scope to build a CanvasPattern). */
function resolveFill(fill: Fill | null): string | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'gradient') {
    return fill.stops.length > 0 ? hexToRgba(fill.stops[0].color) : null;
  }
  // Pattern fills degrade to their foreground colour outside the canvas-aware
  // path; full bitmap rendering happens via resolveShapeFill (core resolveFill).
  if (fill.fillType === 'pattern') return hexToRgba(fill.fg);
  return null;
}

/** Context-aware fill resolver that creates a CanvasGradient for gradient
 * fills and a CanvasPattern for preset pattern fills. */
function resolveShapeFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): string | CanvasGradient | CanvasPattern | null {
  return resolveFillCore(fill, ctx, x, y, w, h);
}

// ===== Text layout helpers =====

/**
 * Draw a text underline at the given baseline. ECMA-376 §21.1.2.3.16
 * defines the OOXML underline enum; we map each value to a Canvas dash
 * pattern + line-weight pair. "wavy*" is approximated by a sine curve
 * traced as a polyline so the glyph stays legibly distinct from "dotted".
 */
function drawUnderline(
  ctx: CanvasRenderingContext2D,
  x: number,
  baseline: number,
  width: number,
  sizePx: number,
  color: string,
  style: string | undefined,
): void {
  const baseLineW = Math.max(1, sizePx * 0.05);
  const heavy = style?.endsWith('Heavy') ?? false;
  const lineW = heavy ? baseLineW * 1.8 : baseLineW;
  const y = baseline + Math.max(2, lineW);
  ctx.strokeStyle = color;
  ctx.lineWidth = lineW;
  ctx.setLineDash([]);

  // Dash patterns scaled by lineW so they stay proportional at any font size.
  // Values mirror prstDash (§20.1.10.49 ST_PresetLineDashVal) where the
  // underline enum reuses the same shape names.
  const dashFor = (s: string): number[] => {
    switch (s) {
      case 'dotted':
      case 'dottedHeavy':         return [1.5, 3];
      case 'dash':
      case 'dashHeavy':           return [6, 3];
      case 'dashLong':
      case 'dashLongHeavy':       return [10, 4];
      case 'dotDash':
      case 'dotDashHeavy':        return [6, 3, 1.5, 3];
      case 'dotDotDash':
      case 'dotDotDashHeavy':     return [6, 3, 1.5, 3, 1.5, 3];
      default:                    return [];
    }
  };

  if (style && style.startsWith('wavy')) {
    // Sine wave with amplitude ≈ lineW and wavelength ≈ 6×lineW.
    const amp = lineW;
    const wavelength = lineW * 6;
    ctx.beginPath();
    ctx.moveTo(x, y);
    const step = Math.max(1, lineW * 0.5);
    for (let dx = 0; dx <= width; dx += step) {
      const yy = y + Math.sin((dx / wavelength) * Math.PI * 2) * amp;
      ctx.lineTo(x + dx, yy);
    }
    ctx.stroke();
    if (style === 'wavyDbl') {
      // Second wave below, offset by 2.5×amp.
      ctx.beginPath();
      ctx.moveTo(x, y + amp * 2.5);
      for (let dx = 0; dx <= width; dx += step) {
        const yy = y + amp * 2.5 + Math.sin((dx / wavelength) * Math.PI * 2) * amp;
        ctx.lineTo(x + dx, yy);
      }
      ctx.stroke();
    }
    return;
  }

  if (style === 'dbl') {
    const offset = lineW * 1.4;
    ctx.beginPath();
    ctx.moveTo(x, y - offset / 2);
    ctx.lineTo(x + width, y - offset / 2);
    ctx.moveTo(x, y + offset / 2);
    ctx.lineTo(x + width, y + offset / 2);
    ctx.stroke();
    return;
  }

  ctx.setLineDash(dashFor(style ?? 'sng').map((v) => v * lineW));
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(x + width, y);
  ctx.stroke();
  ctx.setLineDash([]);
}

// ── Math (OMML) rendering ──────────────────────────────────────────────────
// Equations are converted to SVG by MathJax once, cached by their MathNode[]
// reference (stable from parse through render), then drawn as images inline.
// Mirrors the docx renderer's pipeline so the typesetting is identical.
interface MathRender {
  /** The equation rasterized as opaque black glyphs on transparent. */
  img: HTMLImageElement;
  /** baseline-relative extents in em (1em = the equation's font size in px). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
  /** Per-colour tinted copies (the black `img` recoloured via source-in). */
  tinted: Map<string, CanvasImageSource>;
}
const mathRenders = new WeakMap<MathNode[], MathRender>();

/**
 * Tint the cached black equation image to `color`. PowerPoint equations follow
 * the run/paragraph text colour, which varies per slide (white on dark, theme
 * accents, …), so we recolour the glyphs at draw time via `source-in` instead
 * of baking a single colour. Cached per colour on the render.
 */
function tintedMathImage(render: MathRender, color: string): CanvasImageSource {
  const cached = render.tinted.get(color);
  if (cached) return cached;
  const iw = render.img.naturalWidth || 1;
  const ih = render.img.naturalHeight || 1;
  const canvas = document.createElement('canvas');
  canvas.width = iw;
  canvas.height = ih;
  const cx = canvas.getContext('2d');
  if (!cx) return render.img;
  cx.drawImage(render.img, 0, 0, iw, ih);
  cx.globalCompositeOperation = 'source-in';
  cx.fillStyle = color;
  cx.fillRect(0, 0, iw, ih);
  render.tinted.set(color, canvas);
  return canvas;
}

function svgToImage(svg: string): Promise<HTMLImageElement> {
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  const img = new Image();
  return new Promise((resolve, reject) => {
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = url;
  });
}

// Rasterization resolution for equation SVGs, in px per em. A MathJax SVG
// carries its size in `ex` units, so an `<img>` rasterizes it at a small
// intrinsic size and `drawImage` then upscales it — blurry on HiDPI canvases.
// Forcing an explicit px width/height makes the browser rasterize at this
// resolution instead; 256 px/em stays crisp for equations well past 40pt even
// at devicePixelRatio 3 (40pt → ~53px/em × 3 ≈ 160 px/em needed).
const MATH_RASTER_PX_PER_EM = 256;

/** Pin the SVG root to an explicit high-resolution px size before rasterizing. */
function sizeSvgForRaster(svg: string, widthEm: number, heightEm: number): string {
  const w = Math.max(1, Math.round(widthEm * MATH_RASTER_PX_PER_EM));
  const h = Math.max(1, Math.round(heightEm * MATH_RASTER_PX_PER_EM));
  return svg.replace(/<svg([^>]*?)>/, (_m, attrs: string) => {
    const cleaned = attrs.replace(/\s(?:width|height)="[^"]*"/g, '');
    return `<svg${cleaned} width="${w}" height="${h}">`;
  });
}

/** Gather every math run reachable from a slide's shapes and table cells. */
function collectSlideMathRuns(slide: Slide): { nodes: MathNode[]; display: boolean }[] {
  const found: { nodes: MathNode[]; display: boolean }[] = [];
  const fromBody = (body: TextBody | null) => {
    if (!body) return;
    for (const para of body.paragraphs) {
      for (const run of para.runs) {
        if (run.type === 'math') found.push({ nodes: run.nodes, display: run.display });
      }
    }
  };
  for (const el of slide.elements) {
    if (el.type === 'shape') fromBody(el.textBody);
    else if (el.type === 'table') {
      for (const row of el.rows) for (const cell of row.cells) fromBody(cell.textBody);
    }
  }
  return found;
}

/**
 * Pre-render every equation on the slide to an image (async), so the
 * synchronous layout/draw passes can place them by reading cached extents.
 * A conversion failure leaves the equation unrendered rather than throwing.
 */
export async function prepareSlideMath(slide: Slide, math: MathRenderer): Promise<void> {
  const runs = collectSlideMathRuns(slide);
  if (runs.length === 0) return;
  await math.loadMathJax();
  for (const r of runs) {
    if (mathRenders.has(r.nodes)) continue;
    try {
      const out = await math.mathMLToSvg(mathToMathML(r.nodes, r.display));
      const sized = sizeSvgForRaster(recolorSvg(out.svg, '#000000'), out.widthEm, out.ascentEm + out.descentEm);
      const img = await svgToImage(sized);
      mathRenders.set(r.nodes, {
        img,
        widthEm: out.widthEm,
        ascentEm: out.ascentEm,
        descentEm: out.descentEm,
        tinted: new Map(),
      });
    } catch {
      // leave unrendered
    }
  }
}

type LayoutSegment = {
  text: string;
  font: string;
  sizePx: number;
  color: string;
  underline: boolean;
  /** OOXML rPr @u value when not the default "sng": "dbl"/"dotted"/"wavy"/etc. */
  underlineStyle?: string;
  /** rgba() colour for the underline when uFill overrides the text colour. */
  underlineColor?: string;
  strikethrough: boolean;
  /** Two parallel strike lines (rPr strike="dblStrike"). */
  strikeDouble?: boolean;
  /** Extra inter-character spacing in px (already scaled). */
  letterSpacingPx?: number;
  baseline?: number;
  /** Run-level glyph drop shadow (rPr > effectLst > outerShdw). */
  shadow?: import('@silurus/ooxml-core').Shadow;
  /** Run-level glyph outline (rPr > a:ln). Width in EMU; renderer scales. */
  outline?: import('@silurus/ooxml-core').TextOutline;
  /**
   * Present when this segment is an OMML equation drawn as an image instead of
   * text (`text` is then ""). Box metrics are in px at the current scale.
   */
  math?: {
    nodes: MathNode[];
    display: boolean;
    width: number;
    ascent: number;
    descent: number;
  };
};

interface LayoutLine {
  segments: LayoutSegment[];
  /** Segments right-aligned at a tab stop (set when paragraph contains \t and a right-aligned tabStop) */
  tabStop?: {
    /** Tab stop position in px from the left edge of the text area (bx + lPad + tabStop.px = canvas X) */
    px: number;
    algn: string;
    segments: LayoutSegment[];
  };
}

/**
 * Resolve OOXML theme font references (e.g. "+mn-ea", "+mj-lt") to CSS-safe font names.
 * Canvas will silently ignore an invalid CSS font string, keeping whatever font was set before —
 * leading to wrong text size. Map theme references to generic families as a safe fallback.
 */
const WINGDINGS_MAP: Record<number, string> = {
  0x21: '✏', 0x22: '✂', 0x23: '✁', 0x24: '👁',
  0x4A: '☺', 0x4B: '☻', 0x4C: '☹',
  0x76: '✔', 0xFC: '✓', 0xFB: '✗', 0xFE: '■',
  0xA7: '▪', 0xB7: '•', 0xB8: '◦', 0xB9: '–',
  0xF0A7: '▪', 0xF0B7: '•',
};

function applySymbolFont(char: string, fontFamily: string): string {
  const lower = fontFamily.toLowerCase();
  if (lower.includes('wingdings') || lower === 'symbol') {
    const code = char.charCodeAt(0);
    return WINGDINGS_MAP[code] ?? char;
  }
  return char;
}

function normalizeFontFamily(family: string | null, rc: RenderContext): string {
  if (!family) return rc.themeMinorFont ?? 'sans-serif';
  if (family.startsWith('+')) {
    // +mn-lt = minor Latin, +mj-lt = major Latin, +mn-ea = minor East Asian, +mj-ea = major East Asian
    if (family === '+mj-lt' || family === '+mj-ea' || family === '+mj-cs') {
      return rc.themeMajorFont ?? 'sans-serif';
    }
    // +mn-lt, +mn-ea, +mn-cs, or any other + prefix → minor font
    return rc.themeMinorFont ?? 'sans-serif';
  }
  // OOXML typeface sometimes appends ",<generic>" hint (e.g. "Wingdings,Sans-Serif").
  // Strip it so the CSS font name resolves to the actual named font.
  const primary = family.split(',')[0].trim();
  if (!primary) return rc.themeMinorFont ?? 'sans-serif';
  return primary;
}

/** CSS generic font families — must NOT be quoted in a canvas font string. */
const CSS_GENERIC_FAMILIES = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
]);

/** Infer a CSS generic fallback from a named font so missing fonts degrade consistently. */
function genericFallback(family: string): string {
  const l = family.toLowerCase();
  if (/mono|courier|consolas|等幅|gothic_m/.test(l)) return 'monospace';
  // Serif: mincho (Japanese serif), roman, times, garamond, georgia, palatino, etc.
  if (/mincho|明朝|roman|times|garamond|georgia|palatino|century|didot/.test(l)) return 'serif';
  // Everything else (gothic, kaku, round, rounded, sans, grotesk, …) → sans-serif
  return 'sans-serif';
}

/**
 * Office fonts → metric-compatible, freely-distributable substitutes
 * (`presentation.ts` preloads these webfonts). Putting the substitute in the
 * canvas font stack means a viewer that lacks Calibri/Cambria (macOS, Linux)
 * renders at the SAME advance widths as PowerPoint instead of a wider system
 * serif/sans — without which `wrap="none"` lines overflow the slide. Keys are
 * lower-cased; both the Office face and its substitute share glyph metrics.
 */
const OFFICE_FONT_SUBSTITUTE: Record<string, string> = {
  'calibri': 'Carlito',
  'calibri light': 'Carlito',
  'cambria': 'Caladea',
  'cambria math': 'Caladea',
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL slides (e.g. sample-10, which requests Sakkal Majalla /
  // Univers Next Arabic) render with a real web font instead of an oversized
  // OS fallback. "Naskh" covers traditional serif-like Arabic faces; "Sans"
  // covers the modern geometric ones.
  'sakkal majalla': 'Noto Naskh Arabic',
  'traditional arabic': 'Noto Naskh Arabic',
  'simplified arabic': 'Noto Naskh Arabic',
  'arabic typesetting': 'Noto Naskh Arabic',
  'univers next arabic': 'Noto Sans Arabic',
};

/** Generic Arabic fallbacks appended to an Arabic-script font's canvas stack
 *  (before the CSS generic) so Arabic glyphs in an Arabic-targeted family that
 *  the host lacks still resolve to a real Arabic web font when `useGoogleFonts`
 *  is on. */
const ARABIC_FALLBACKS = '"Noto Naskh Arabic", "Noto Sans Arabic"';

/**
 * True when `family` names an Arabic-script face. Only such faces get the Noto
 * Arabic web fonts appended to their canvas stack.
 *
 * These fallback faces (esp. Noto Naskh Arabic) also carry serif-style *Latin*
 * glyphs, so appending them to every font stack made Latin text in an
 * uninstalled Latin/CJK face (e.g. a Japanese gothic carrying "About us") fall
 * into the serif Naskh face instead of degrading to the sans-serif generic.
 * Gating on Arabic-script faces keeps RTL decks (Amiri, Sakkal Majalla, …)
 * correct while letting Latin/CJK text degrade to the right generic.
 */
function isArabicScriptFace(family: string): boolean {
  // Faces we explicitly substitute to a Noto Arabic web font are Arabic-script.
  if (OFFICE_FONT_SUBSTITUTE[family.toLowerCase()]?.includes('Arabic')) return true;
  const l = family.toLowerCase();
  // Common Arabic-script family-name markers (Latin transliterations + Arabic).
  return /arabic|naskh|kufi|nastaliq|amiri|scheherazade|lateef|aldhabi|urdu|farsi|العرب|[؀-ۿ]/.test(l);
}

function buildFont(bold: boolean, italic: boolean, sizePx: number, family: string, rc: RenderContext): string {
  const style  = italic ? 'italic ' : '';
  const weight = bold   ? 'bold '   : '';
  const normalized = normalizeFontFamily(family, rc);
  if (CSS_GENERIC_FAMILIES.has(normalized)) {
    return `${style}${weight}${sizePx}px ${normalized}`;
  }
  // Named font + (metric-compatible substitute, if an Office face) + Arabic web
  // fonts (only for Arabic-script faces — see isArabicScriptFace) + inferred
  // generic fallback, so browsers degrade consistently when the exact typeface
  // is not installed.
  const sub = OFFICE_FONT_SUBSTITUTE[normalized.toLowerCase()];
  const subPart = sub ? `"${sub}", ` : '';
  const arabicPart = isArabicScriptFace(normalized) ? `${ARABIC_FALLBACKS}, ` : '';
  return `${style}${weight}${sizePx}px "${normalized}", ${subPart}${arabicPart}${genericFallback(normalized)}`;
}

/**
 * Lay out a paragraph into display lines.
 * Handles:
 *  - Explicit line breaks (TextRun type='break')
 *  - Space-based word wrap (Latin text)
 *  - Character-level wrap fallback for CJK / words wider than container
 *  - Tab stops (right-aligned and left-aligned)
 *
 * @param marLPx  Paragraph left margin in canvas px (used for tab stop position calculation)
 */
/**
 * Decide whether `<a:spAutoFit/>` should let text wrap based on the paragraphs'
 * natural single-line width. Returns true when at least one paragraph would
 * exceed the available text width if rendered without wrapping — that's the
 * cue PowerPoint uses to switch from "grow shape horizontally" to "wrap and
 * grow shape vertically" mode (ECMA-376 §20.1.10.5).
 *
 * The measurement is intentionally conservative: it sums all run widths in a
 * paragraph at the run's own font size, ignoring tab stops and explicit
 * `\t` runs. That matches what PowerPoint compares — "one continuous line
 * of glyphs" — and avoids over-eager wrap when a paragraph is barely wider
 * than the bbox due to font-measurement drift.
 */
function naturalWidthExceedsBbox(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bw: number,
  lPad: number,
  rPad: number,
  scale: number,
  rc: RenderContext,
): boolean {
  const bodyDefaultFontPx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale;
  for (const para of body.paragraphs) {
    const marLPx = emuToPx(para.marL, scale);
    const marRPx = emuToPx(para.marR, scale);
    const indentPx = emuToPx(para.indent, scale);
    // First-line indent eats into the available width when positive; a
    // negative indent (hanging) is the bullet's gutter and doesn't reduce
    // the usable text room.
    const firstLineIndent = Math.max(0, indentPx);
    const textMaxW = bw - lPad - rPad - marLPx - marRPx - firstLineIndent;
    let lineW = 0;
    for (const run of para.runs) {
      if (run.type !== 'text') continue;
      const sizePx = run.fontSize != null
        ? run.fontSize * PT_TO_EMU * scale
        : (para.defFontSize != null
            ? para.defFontSize * PT_TO_EMU * scale
            : bodyDefaultFontPx);
      const family = normalizeFontFamily(run.fontFamily ?? para.defFontFamily ?? null, rc);
      const isBold = run.bold ?? para.defBold ?? body.defaultBold ?? false;
      const isItalic = run.italic ?? para.defItalic ?? body.defaultItalic ?? false;
      ctx.font = buildFont(isBold, isItalic, sizePx, family, rc);
      lineW += ctx.measureText(run.text).width;
      if (lineW > textMaxW) return true;
    }
  }
  return false;
}

function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  para: Paragraph,
  maxWidthPx: number,
  defaultFontSizePx: number,
  defaultColor: string,
  scale: number,
  marLPx: number,
  defaultBold: boolean = false,
  defaultItalic: boolean = false,
  fontScale: number = 1.0,
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null },
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: LayoutLine = { segments: [] };
  let lineW = 0; // current line's accumulated width
  // ECMA-376 §17.18.93 ST_TextWrappingType "square" is whitespace-aware: a
  // non-whitespace token is never broken away from the preceding non-whitespace
  // content. We only allow a wrap before a token if at least one whitespace
  // run has appeared on the current line — otherwise the line overflows the
  // shape (PowerPoint's actual behavior, e.g. "YoY+11.9%" mixed-size runs in
  // sample-2 slide-7 stay on one line even though the bbox is tight).
  let hasWhitespaceOnLine = false;

  // Tab stop state: once we hit a \t we switch to collecting tabStop.segments
  let tabActive = false;
  let tabStopPx = 0;   // position of tab stop from text area left (px)

  const newLine = () => {
    lines.push(currentLine);
    currentLine = { segments: [] };
    lineW = 0;
    tabActive = false; // reset tab state per line
    hasWhitespaceOnLine = false;
  };

  // Push to the active segment list (main or tab-stop group)
  const push = (
    text: string,
    font: string,
    sizePx: number,
    color: string,
    underline: boolean,
    strikethrough: boolean,
    baseline?: number,
    extras?: {
      strikeDouble?: boolean;
      letterSpacingPx?: number;
      underlineStyle?: string;
      underlineColor?: string;
      shadow?: import('@silurus/ooxml-core').Shadow;
      outline?: import('@silurus/ooxml-core').TextOutline;
    },
  ) => {
    if (!text) return;
    ctx.font = font;
    const lsPx = extras?.letterSpacingPx ?? 0;
    const baseW = ctx.measureText(text).width;
    // Letter spacing adds an extra gap between every character, including
    // after the last one — matches the "advance width" semantics of OOXML
    // spc (each glyph's advance grows by spc points). Tab stops measure the
    // same way so this stays consistent with measureText below.
    const w = baseW + lsPx * text.length;
    const strikeDouble = extras?.strikeDouble;
    const underlineStyle = extras?.underlineStyle;
    const underlineColor = extras?.underlineColor;
    const shadow = extras?.shadow;
    const outline = extras?.outline;
    // Shadow / outline use object identity for merging — adjacent runs share
    // the same object since the run is parsed once. Different objects (or
    // one set / one missing) force a new segment.
    const sameMeta = (a: LayoutSegment) =>
      !a.math &&
      a.font === font &&
      a.color === color &&
      a.underline === underline &&
      (a.underlineStyle ?? '') === (underlineStyle ?? '') &&
      (a.underlineColor ?? '') === (underlineColor ?? '') &&
      a.strikethrough === strikethrough &&
      (a.strikeDouble ?? false) === (strikeDouble ?? false) &&
      (a.letterSpacingPx ?? 0) === lsPx &&
      a.baseline === baseline &&
      a.shadow === shadow &&
      a.outline === outline;
    if (tabActive && currentLine.tabStop) {
      const segs = currentLine.tabStop.segments;
      const last = segs.at(-1);
      if (last && sameMeta(last)) {
        last.text += text;
      } else {
        segs.push({ text, font, sizePx, color, underline, underlineStyle, underlineColor, strikethrough, strikeDouble, letterSpacingPx: lsPx || undefined, baseline, shadow, outline });
      }
    } else {
      lineW += w;
      const last = currentLine.segments.at(-1);
      if (last && sameMeta(last)) {
        last.text += text;
      } else {
        currentLine.segments.push({ text, font, sizePx, color, underline, underlineStyle, underlineColor, strikethrough, strikeDouble, letterSpacingPx: lsPx || undefined, baseline, shadow, outline });
      }
    }
  };

  for (const run of para.runs) {
    if (run.type === 'break') {
      newLine();
      continue;
    }

    // ── OMML equation ─────────────────────────────────────────────────────
    if (run.type === 'math') {
      const render = mathRenders.get(run.nodes);
      // Equation font size: explicit run size (pt→px) else paragraph default.
      const emPx = run.fontSize != null
        ? run.fontSize * PT_TO_EMU * scale * fontScale
        : defaultFontSizePx;
      const width = render ? render.widthEm * emPx : 0;
      const ascent = render ? render.ascentEm * emPx : 0;
      const descent = render ? render.descentEm * emPx : 0;
      // Block (display) math gets its own line; the draw pass centres it.
      if (run.display && lineW > 0) newLine();
      else if (lineW + width > maxWidthPx && lineW > 0) newLine();
      currentLine.segments.push({
        text: '',
        font: `${emPx}px sans-serif`,
        sizePx: emPx,
        // Equations follow their own run colour (e.g. a purple title); the
        // draw pass tints the glyph image to this colour. Fall back to the
        // paragraph/body default when the run carries no explicit colour.
        color: run.color ? hexToRgba(run.color) : defaultColor,
        underline: false,
        strikethrough: false,
        math: { nodes: run.nodes, display: run.display, width, ascent, descent },
      });
      lineW += width;
      if (run.display) newLine();
      continue;
    }

    const sizePx = run.fontSize != null ? run.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
    // Font family cascade: run → paragraph defFontFamily → theme minor font → 'sans-serif'
    const family = normalizeFontFamily(run.fontFamily ?? para.defFontFamily ?? null, rc);
    // East Asian font (rPr > ea) — used for CJK glyphs when set; otherwise
    // CJK characters reuse the latin font. ECMA-376 §21.1.2.3.7.
    const familyEa = run.fontFamilyEa
      ? normalizeFontFamily(run.fontFamilyEa, rc)
      : null;
    // Hyperlink runs without an explicit colour pick up the theme hlink colour
    // (ECMA-376 §20.1.2.3.5 — hyperlinks inherit theme hyperlink slot).
    let color: string;
    if (run.color) {
      color = hexToRgba(run.color);
    } else if (run.hyperlink && rc.themeHlinkColor) {
      color = hexToRgba(rc.themeHlinkColor);
    } else {
      color = defaultColor;
    }
    // Cascade: run → paragraph defRPr → body/layout default → false
    const isBold   = run.bold   ?? para.defBold   ?? defaultBold;
    const isItalic = run.italic ?? para.defItalic ?? defaultItalic;
    const font   = buildFont(isBold, isItalic, sizePx, family, rc);
    const fontEa = familyEa
      ? buildFont(isBold, isItalic, sizePx, familyEa, rc)
      : font;
    ctx.font = font;

    // ECMA-376 §21.1.2.3.13 — caps transforms the rendered glyphs without
    // changing the underlying text. "small" emulated as upper-case glyphs at
    // ~80% size is the long-established Office fallback when the font lacks
    // smcp; we just upper-case for now and rely on the configured size.
    const caps = run.caps;
    let baseText = run.text;
    if (caps === 'all' || caps === 'small') baseText = baseText.toUpperCase();

    // Resolve field values (e.g. slidenum → actual slide number)
    const runText = (run.fieldType === 'slidenum' && slideNumber !== undefined)
      ? String(slideNumber)
      : baseText;

    // Hyperlink runs render underlined unless an explicit u attribute already
    // says otherwise. Spec: ECMA-376 §20.1.2.3.5 (hyperlinks default to the
    // hlink character style, which underlines).
    const segUnderline = run.underline || (run.hyperlink !== undefined);
    const segStrikeDouble = run.strikeDouble === true;
    // letterSpacing arrives in points; convert to canvas px using the same
    // EMU→px scale the renderer applies to font sizes.
    const lsPx = run.letterSpacing != null ? run.letterSpacing * PT_TO_EMU * scale : 0;
    const segExtras = {
      strikeDouble: segStrikeDouble,
      letterSpacingPx: lsPx,
      underlineStyle: run.underlineStyle,
      underlineColor: run.underlineColor ? hexToRgba(run.underlineColor) : undefined,
      shadow: run.shadow,
      outline: run.outline,
    };

    // Split on whitespace boundaries, keeping the whitespace tokens
    const tokens = runText.split(/(\s+)/);

    for (const token of tokens) {
      if (!token) continue;

      // ── Tab character ────────────────────────────────────────────────────
      if (/^\t+$/.test(token)) {
        // Find first tab stop whose position (from text area left) is beyond the current pen
        const currentAbsW = marLPx + lineW; // current position from text area left
        const ts = (para.tabStops ?? []).find(
          (t: TabStop) => emuToPx(t.pos, scale) > currentAbsW
        );
        if (ts) {
          tabStopPx = emuToPx(ts.pos, scale);
          if (ts.algn === 'r' || ts.algn === 'ctr') {
            // Switch to tab-stop accumulation mode
            tabActive = true;
            currentLine.tabStop = { px: tabStopPx, algn: ts.algn, segments: [] };
          } else {
            // Left-aligned tab: advance lineW to the tab stop
            lineW = tabStopPx - marLPx;
          }
        } else {
          // No matching tab stop — treat as a single space
          push(' ', font, sizePx, color, segUnderline, run.strikethrough, undefined, segExtras);
        }
        continue;
      }

      ctx.font = font;
      const tokW = ctx.measureText(token).width;
      const isWhitespace = /^\s+$/.test(token);

      // If already in tab mode, collect all text into tabStop.segments (no wrap)
      if (tabActive) {
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        continue;
      }

      // CJK characters allow line-breaking at any character boundary (no whitespace
      // needed). When a token contains CJK, wrap character-by-character so that CJK
      // text flows onto the same line as preceding Latin text (e.g. "EC市場で…").
      // Per-character font dispatch picks `fontEa` for CJK glyphs when the run
      // declared an explicit East Asian typeface (rPr > ea); other characters
      // keep the Latin font so the latin/ea boundary mid-token stays clean.
      const CJK_RE = /[\u3000-\u9FFF\uAC00-\uD7FF\uF900-\uFAFF\uFF00-\uFFEF]/;
      const hasCJK = CJK_RE.test(token);
      if (hasCJK) {
        for (const ch of token) {
          const chFont = CJK_RE.test(ch) ? fontEa : font;
          ctx.font = chFont;
          const chW = ctx.measureText(ch).width;
          if (lineW + chW > maxWidthPx && lineW > 0) newLine();
          push(ch, chFont, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        }
        continue;
      }

      if (lineW + tokW <= maxWidthPx) {
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        if (isWhitespace) hasWhitespaceOnLine = true;
      } else if (isWhitespace) {
        if (lineW > 0) newLine();
      } else if (tokW > maxWidthPx) {
        if (lineW > 0) newLine();
        for (const ch of token) {
          ctx.font = font;
          const chW = ctx.measureText(ch).width;
          if (lineW + chW > maxWidthPx && lineW > 0) newLine();
          push(ch, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        }
      } else if (!hasWhitespaceOnLine) {
        // No whitespace yet on this line — wrapping here would tear an
        // unbroken sequence of non-whitespace text (e.g. "YoY+11.9%" split
        // across mixed-size runs). Office never breaks mid-sequence in that
        // case; it lets the shape overflow and relies on spAutoFit / lIns to
        // size the bbox correctly. Match that behavior.
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
      } else {
        newLine();
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
      }
    }
  }

  // Always emit the last (possibly empty) line
  lines.push(currentLine);

  return lines;
}

// ===== Element renderers =====

function renderBackground(
  ctx: CanvasRenderingContext2D,
  fill: Fill | null,
  canvasW: number,
  canvasH: number
) {
  const bg = resolveShapeFill(fill, ctx, 0, 0, canvasW, canvasH);
  ctx.fillStyle = bg ?? '#FFFFFF';
  ctx.fillRect(0, 0, canvasW, canvasH);
}

function applyShadow(ctx: CanvasRenderingContext2D, shadow: Shadow | null, scale: number) {
  if (!shadow) return;
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dist = emuToPx(shadow.dist, scale);
  ctx.shadowColor = hexToRgba(shadow.color, shadow.alpha);
  ctx.shadowBlur = emuToPx(shadow.blur, scale);
  ctx.shadowOffsetX = Math.cos(dirRad) * dist;
  ctx.shadowOffsetY = Math.sin(dirRad) * dist;
}

/**
 * ECMA-376 §20.1.8.17 — apply glow as a coloured halo. The Canvas shadow
 * primitive is a fine fit: zero offset + the glow's blur radius produces a
 * symmetric coloured blur centred on every drawn pixel.
 */
function applyGlow(ctx: CanvasRenderingContext2D, glow: Glow | null | undefined, scale: number) {
  if (!glow) return;
  ctx.shadowColor = hexToRgba(glow.color, glow.alpha);
  ctx.shadowBlur = emuToPx(glow.radius, scale);
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

function clearShadow(ctx: CanvasRenderingContext2D) {
  ctx.shadowColor = 'transparent';
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

/**
 * Preset-geometry text rectangle (ECMA-376 §20.1.9.21 `<a:rect>` /
 * presetTextRectangle), as an absolute sub-rect of the shape bbox. PowerPoint
 * lays text out inside this rect (then applies the `lIns/tIns/rIns/bIns`
 * insets), NOT in the full bounding box. For arrows the text rect is the shaft
 * (left of the arrowhead), so centered text sits in the body — without this a
 * `rightArrow`'s label drifts right into the arrowhead. Returns null when the
 * geometry's text rect is the whole bbox (the common case).
 *
 * Arrowhead length uses `ss = min(w, h)` per the spec's `dx1 = ss * adj2`.
 */
function presetTextRect(
  geom: string,
  x: number, y: number, w: number, h: number,
  adj1?: number | null, adj2?: number | null,
): { tx: number; ty: number; tw: number; th: number } | null {
  const ss = Math.min(w, h);
  switch (geom) {
    case 'rightarrow':
    case 'leftarrow': {
      const a1 = Math.min(Math.max(adj1 ?? 50000, 0), 100000); // shaft height fraction
      const a2 = Math.min(Math.max(adj2 ?? 50000, 0), 100000); // arrowhead length fraction
      const dx = (ss * a2) / 100000;        // arrowhead length (ss-based, ECMA dx1)
      const dy = (h * a1) / 200000;          // shaft half-height about the center
      const ty = y + h / 2 - dy;
      const th = 2 * dy;
      const tw = Math.max(0, w - dx);
      return geom === 'rightarrow'
        ? { tx: x, ty, tw, th }              // shaft on the left, arrowhead on the right
        : { tx: x + dx, ty, tw, th };        // leftarrow: arrowhead on the left
    }
    case 'roundrect': {
      // Text rect inset from the rounded corners: il = radius * (1 - 1/√2),
      // radius = ss * adj / 100000 (ECMA roundRect adj default 16667).
      const a = Math.min(Math.max(adj1 ?? 16667, 0), 100000) as number;
      const il = (ss * a) / 100000 * (1 - 1 / Math.SQRT2);
      return { tx: x + il, ty: y + il, tw: Math.max(0, w - 2 * il), th: Math.max(0, h - 2 * il) };
    }
    default:
      return null;
  }
}

function renderShape(ctx: CanvasRenderingContext2D, el: ShapeElement, scale: number, themeDefaultColor = '#000000', slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null }, onTextRun?: TextRunCallback) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  // anchor="b" + h=0: shape grows upward from y; render stroke as bottom border,
  // then let renderTextBody handle positioning.
  if (h === 0 && el.textBody?.verticalAnchor === 'b') {
    if (el.stroke) {
      ctx.save();
      applyStroke(ctx, el.stroke, scale);
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.stroke();
      ctx.restore();
    }
    if (el.textBody) {
      const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
      renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor, el.rotation, el.flipH, el.flipV, themeDefaultColor, slideNumber, rc, onTextRun);
    }
    return;
  }

  ctx.save();
  if (el.rotation !== 0 || el.flipH || el.flipV) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((el.rotation * Math.PI) / 180);
    if (el.flipH) ctx.scale(-1, 1);
    if (el.flipV) ctx.scale(1, -1);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }

  const geom = el.geometry.toLowerCase();
  const fillStyle = resolveShapeFill(el.fill, ctx, x, y, w, h);

  // Apply shadow before fill/stroke drawing; ctx.restore() will clear it.
  // The Canvas API exposes a single shadow slot, so when both an outer shadow
  // and a glow are configured we let the outer shadow win (visually dominant)
  // and fall back to the glow only when no outer shadow is present. This is
  // a common — and conservative — interpretation of layered effectLst.
  applyShadow(ctx, el.shadow ?? null, scale);
  if (!el.shadow) applyGlow(ctx, el.glow ?? null, scale);

  const CONNECTOR_GEOMS = new Set([
    'line', 'straightconnector1',
    'bentconnector2', 'bentconnector3', 'bentconnector4', 'bentconnector5',
    'curvedconnector2', 'curvedconnector3', 'curvedconnector4', 'curvedconnector5',
  ]);

  // callout1 family — `<a:ln><a:headEnd|tailEnd>` decorate the line path
  // (path 1 in presets.json), not the surrounding text rectangle (path 0).
  const CALLOUT1_GEOMS = new Set([
    'callout1', 'bordercallout1', 'accentcallout1', 'accentbordercallout1',
  ]);

  // ── Dispatch to preset engine when possible ────────────────────────────
  // Preference order: custGeom → generic preset engine → legacy switch.
  // `arc` keeps its bespoke case because its fill semantics (pie-wedge vs
  // open arc) depend on stroke state in ways the engine doesn't express.
  const usePresetEngine =
    !el.custGeom && geom !== 'arc' && hasPreset(geom);

  /**
   * Paint the shape's body (fill + stroke) into an arbitrary target context,
   * dispatching through the same preset / custGeom / legacy paths as the live
   * render. Extracted into a closure so the edge/blur effect helpers
   * (innerShdw §20.1.8.40, softEdge §20.1.8.53, reflection §20.1.8.50) can
   * re-paint the shape onto auxiliary canvases.
   *
   * `silhouette`: when set, paint a flat opaque silhouette in that colour
   * (filled path only, no stroke, no gradients) — used by innerShdw to build
   * its mask. The silhouette honours the same even-odd / preset path topology.
   */
  const paintShapeBody = (
    target: CanvasRenderingContext2D,
    silhouette?: string,
  ): void => {
    const tFill = silhouette ?? fillStyle;
    const tStroke = silhouette
      ? null
      : el.stroke
        ? () => {
            applyStroke(target, el.stroke!, scale);
            target.stroke();
          }
        : null;
    const tClearShadow = () => clearShadow(target);

    if (usePresetEngine && !silhouette) {
      renderPresetShape(
        target, geom, x, y, w, h,
        [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8],
        tFill, tStroke, tClearShadow,
      );
      return;
    }

    target.beginPath();
    if (el.custGeom && el.custGeom.length > 0) {
      buildCustomPath(target, el.custGeom, x, y, w, h);
    } else if (usePresetEngine) {
      // Silhouette of a preset shape: build a single filled outline. Preset
      // path 0 is the body outline (secondary paths are highlights), so the
      // legacy buildShapePath gives a faithful silhouette of the body.
      buildShapePath(target, geom, x, y, w, h, el.adj, el.adj2, el.adj3, el.adj4);
    } else {
      buildShapePath(target, geom, x, y, w, h, el.adj, el.adj2, el.adj3, el.adj4);
    }
    if (tFill && geom !== 'arc') {
      target.fillStyle = tFill;
      if (geom === 'donut' || geom === 'smileyface' || geom === 'frame') {
        target.fill('evenodd');
      } else {
        target.fill();
      }
      if (!silhouette) tClearShadow();
    }
    if (tStroke) {
      tStroke();
    }
  };

  // ── effectLst edge/blur effects (independent siblings, ECMA-376 §20.1.8.25)
  // Device-pixel canvas extent for the auxiliary effect canvases.
  const deviceW = (ctx.canvas as { width: number }).width || 0;
  const deviceH = (ctx.canvas as { height: number }).height || 0;
  // The effect helpers operate in DEVICE pixels: the aux silhouette is painted
  // through the live transform (which already folds in devicePixelRatio +
  // rotation + flip), and the blit happens at identity. So bbox / radii passed
  // to the helpers must be in device px too. The uniform device scale equals
  // sqrt(|det|) of the transform's linear part — for scale(dpr)·rot·flip this
  // is exactly dpr, independent of rotation or mirroring.
  const liveTransform = ctx.getTransform();
  const det = Math.abs(liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;
  const effBBox = { x: x * devScale, y: y * devScale, w: w * devScale, h: h * devScale };
  const effScale = scale * devScale; // EMU → device px
  const applyLiveTransform = (c: CanvasRenderingContext2D) => {
    c.setTransform(liveTransform);
  };

  // Reflection sits BEHIND/below the shape — draw it first so the body paints
  // on top. §20.1.8.50. The aux silhouette bakes in the live rotation/flip via
  // setTransform, and the helper's mirror transform operates in device space,
  // so the live ctx must blit at identity.
  if (el.reflection && deviceW > 0 && deviceH > 0) {
    ctx.save();
    ctx.setTransform(new DOMMatrix());
    applyReflection(
      ctx, (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D); },
      effBBox, el.reflection, effScale, deviceW, deviceH,
    );
    ctx.restore();
  }

  // softEdge feathers the whole body, so it REPLACES the direct body paint.
  // §20.1.8.53. When absent, paint the body normally.
  if (el.softEdge && deviceW > 0 && deviceH > 0) {
    // The feathered body is composited in untransformed device space, so reset
    // the live transform around the blit and re-apply the same transform when
    // painting the body into the aux canvas.
    ctx.save();
    ctx.setTransform(new DOMMatrix());
    applySoftEdge(
      ctx, (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D); },
      effBBox, el.softEdge, effScale, deviceW, deviceH,
      // Mask is the flat filled silhouette (no stroke) — see applySoftEdge.
      (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D, '#000'); },
    );
    ctx.restore();
  } else {
    paintShapeBody(ctx);
  }

  // innerShdw casts inward, ON TOP of the fill. §20.1.8.40. Composite after the
  // body. The silhouette callback paints a flat opaque mask.
  if (el.innerShadow && deviceW > 0 && deviceH > 0) {
    ctx.save();
    ctx.setTransform(new DOMMatrix());
    applyInnerShadow(
      ctx, (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D, '#000'); },
      effBBox, el.innerShadow, effScale, deviceW, deviceH,
    );
    ctx.restore();
  }

  if (el.stroke && CONNECTOR_GEOMS.has(geom)) {
    const anchors = getConnectorAnchors(geom, x, y, w, h, [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8]);
    if (anchors) {
      // ECMA-376 §20.1.8.42 — compound line styles. For straight lines /
      // connectors we re-stroke the segment with multiple parallel lines
      // along the perpendicular of the line direction. Curved connectors
      // fall through to the single-stroke fast path (parallel curves are
      // a non-trivial geometric operation).
      const cmpd = el.stroke.cmpd;
      const isStraight = geom === 'line' || geom === 'straightconnector1';
      if (cmpd && isStraight) {
        drawCompoundLine(ctx, anchors.start, anchors.end, el.stroke, cmpd, scale);
      }
      if (el.stroke.tailEnd) {
        drawArrowHead(ctx, anchors.end.x, anchors.end.y, anchors.end.angle, el.stroke.tailEnd, el.stroke, scale);
      }
      if (el.stroke.headEnd) {
        drawArrowHead(ctx, anchors.start.x, anchors.start.y, anchors.start.angle, el.stroke.headEnd, el.stroke, scale);
      }
    }
  } else if (el.stroke && CALLOUT1_GEOMS.has(geom)) {
    // Callout1 family carries an `<a:ln>` whose head/tail decorations belong
    // on the *line* (path 1: m,x1,y1 → l,x2,y2), not on the rectangle (path 0).
    // ECMA-376 callout1 gd: y1=h·adj1/100000, x1=w·adj2/100000 (attach point);
    // y2=h·adj3/100000, x2=w·adj4/100000 (tip). Tip and attach may sit
    // outside the bbox. Compute both, then orient the head/tail along the
    // attach→tip direction.
    const attXf = ((el.adj2 ?? -8333) as number) / 100000;
    const attYf = ((el.adj  ?? 18750) as number) / 100000;
    const tipXf = ((el.adj4 ?? -38333) as number) / 100000;
    const tipYf = ((el.adj3 ?? 112500) as number) / 100000;
    const attX = x + attXf * w;
    const attY = y + attYf * h;
    const tipX = x + tipXf * w;
    const tipY = y + tipYf * h;
    const tailAngle = Math.atan2(tipY - attY, tipX - attX);
    const headAngle = tailAngle + Math.PI;
    if (el.stroke.tailEnd) {
      drawArrowHead(ctx, tipX, tipY, tailAngle, el.stroke.tailEnd, el.stroke, scale);
    }
    if (el.stroke.headEnd) {
      drawArrowHead(ctx, attX, attY, headAngle, el.stroke.headEnd, el.stroke, scale);
    }
  }

  // Render text inside the rotation context so text follows shape rotation
  if (el.textBody) {
    const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
    ctx.save();
    if (el.flipH || el.flipV) {
      const cx = x + w / 2;
      const cy = y + h / 2;
      // The shape itself stays mirrored, but text should remain readable.
      // Apply the same flip again around the shape centre to cancel only the text mirror.
      ctx.translate(cx, cy);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-cx, -cy);
    }
    // For ellipses, PowerPoint positions text relative to the inscribed rectangle
    // (the maximum-area rectangle that fits inside the ellipse: sides = a/√2, b/√2).
    // This only affects non-ctr anchors; ctr anchor is invariant to this inset.
    let tx = x, ty = y, tw = w, th = h;
    if (el.textRect) {
      // SmartArt drawings carry an explicit text frame (<dsp:txXfrm>) that
      // PowerPoint pre-computed against the actual layout — e.g. an arrow's
      // label sits past an overlapping circle node, a roundRect's label avoids
      // an overlapping bottom badge. Honour it verbatim (insets apply within).
      tx = emuToPx(el.textRect.x, scale);
      ty = emuToPx(el.textRect.y, scale);
      tw = emuToPx(el.textRect.width, scale);
      th = emuToPx(el.textRect.height, scale);
    } else if (geom === 'ellipse') {
      const insetX = w * (1 - 1 / Math.SQRT2) / 2;
      const insetY = h * (1 - 1 / Math.SQRT2) / 2;
      tx = x + insetX; ty = y + insetY;
      tw = w / Math.SQRT2; th = h / Math.SQRT2;
    } else {
      // Preset text rectangle (ECMA-376 §20.1.9.21): e.g. an arrow's label sits
      // in the shaft, not the full bbox. Insets (lIns/…) apply within this rect.
      const tr = presetTextRect(geom, x, y, w, h, el.adj, el.adj2);
      if (tr) { tx = tr.tx; ty = tr.ty; tw = tr.tw; th = tr.th; }
    }
    // Pass el.rotation so the text-layer overlay can CSS-rotate the shape div to match.
    renderTextBody(ctx, el.textBody, tx, ty, tw, th, scale, defaultTextColor, el.rotation, false, false, themeDefaultColor, slideNumber, rc, onTextRun);
    ctx.restore();
  }

  ctx.restore();
}

/**
 * Build a canvas path from custGeom path commands.
 * Coordinates are in [0,1] relative to the shape bounding box;
 * the renderer maps them to canvas pixels.
 * Tracks pen position so arcTo can compute the ellipse centre correctly.
 */
const buildCustomPath = buildCustomPathCore;

/**
 * Format an autoNum bullet label from a counter value and OOXML numType.
 * Spec: ECMA-376 §20.1.10.61 (ST_TextAutonumberScheme).
 *
 * Supports the symmetric Plain / Period / ParenR / ParenBoth variants for
 * the four core scripts (arabic, alphaLc/Uc, romanLc/Uc) plus arabicDb
 * (full-width Arabic digits). CJK / Thai / Hebrew / Arabic-Abjad schemes
 * are intentionally not handled here — they fall through to the default
 * `arabicPeriod` rendering rather than emitting a wrong-looking glyph.
 */
function formatAutoNum(counter: number, numType: string): string {
  const arabic = `${counter}`;
  const alphaLc = counter >= 1 && counter <= 26
    ? String.fromCharCode(96 + counter)
    : arabic;
  const alphaUc = counter >= 1 && counter <= 26
    ? String.fromCharCode(64 + counter)
    : arabic;
  const romanLc = toRoman(counter).toLowerCase();
  const romanUc = toRoman(counter);
  // ECMA-376 §20.1.10.61 lists `arabicDb*` as full-width Arabic digits
  // (U+FF10–U+FF19), used in East Asian numbered lists.
  const arabicDb = arabic.replace(/[0-9]/g, (d) =>
    String.fromCharCode(0xff10 + (d.charCodeAt(0) - 0x30)),
  );

  switch (numType) {
    case 'arabicPlain':       return arabic;
    case 'arabicPeriod':      return `${arabic}.`;
    case 'arabicParenR':      return `${arabic})`;
    case 'arabicParenBoth':   return `(${arabic})`;
    case 'arabicDbPlain':     return arabicDb;
    case 'arabicDbPeriod':    return `${arabicDb}.`;
    case 'alphaLcPlain':      return alphaLc;
    case 'alphaLcPeriod':     return `${alphaLc}.`;
    case 'alphaLcParenR':     return `${alphaLc})`;
    case 'alphaLcParenBoth':  return `(${alphaLc})`;
    case 'alphaUcPlain':      return alphaUc;
    case 'alphaUcPeriod':     return `${alphaUc}.`;
    case 'alphaUcParenR':     return `${alphaUc})`;
    case 'alphaUcParenBoth':  return `(${alphaUc})`;
    case 'romanLcPlain':      return romanLc;
    case 'romanLcPeriod':     return `${romanLc}.`;
    case 'romanLcParenR':     return `${romanLc})`;
    case 'romanLcParenBoth':  return `(${romanLc})`;
    case 'romanUcPlain':      return romanUc;
    case 'romanUcPeriod':     return `${romanUc}.`;
    case 'romanUcParenR':     return `${romanUc})`;
    case 'romanUcParenBoth':  return `(${romanUc})`;
    default:                  return `${arabic}.`;
  }
}

function toRoman(n: number): string {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
  }
  return result;
}

function renderTextBody(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  shapeDefaultTextColor: string | null = null,
  shapeRotation = 0,
  shapeFlipH = false,
  shapeFlipV = false,
  themeDefaultColor = '#000000',
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null },
  onTextRun?: TextRunCallback,
  measureOnly = false,
): number | void {
  // Vertical text: rotate rendering context so text flows top-to-bottom.
  // "vert" and "eaVert" both approximate to 90° clockwise rotation.
  // "vert270" rotates 270° (= 90° counterclockwise).
  const isVert    = body.vert === 'vert' || body.vert === 'eaVert';
  const isVert270 = body.vert === 'vert270';

  if (isVert || isVert270) {
    // Set up a rotated coordinate space:
    // Centre of the bounding box remains fixed; swap w and h for the text layout.
    const cx = bx + bw / 2;
    const cy = by + bh / 2;
    const vertRot = isVert ? 90 : -90;

    // Wrap onTextRun to convert from the rotated sub-frame back to the original
    // shape frame so that _buildTextLayer can apply a single CSS rotation.
    //
    // In the recursive call the origin is (-bh/2, -bw/2) with axes (bh, bw).
    // For a run at canvas (penX, cursorY) in that sub-frame:
    //   inShapeX_rec = penX + bh/2,  inShapeY_rec = cursorY + bw/2
    //
    // We need the position in the *original* shape frame so that after
    // CSS rotate(shapeRotation + vertRot) the span lands on the same pixel:
    //   inShapeX_span = penX + bw/2 = inShapeX_rec - bh/2 + bw/2
    //   inShapeY_span = cursorY + bh/2 = inShapeY_rec - bw/2 + bh/2
    const wrappedOnTextRun: TextRunCallback | undefined = onTextRun
      ? (run) => onTextRun({
          ...run,
          inShapeX: run.inShapeX - bh / 2 + bw / 2,
          inShapeY: run.inShapeY - bw / 2 + bh / 2,
          shapeX: bx,
          shapeY: by,
          shapeW: bw,
          shapeH: bh,
          rotation: shapeRotation,
          textBodyRotation: vertRot,
        })
      : undefined;

    if (measureOnly) {
      // The rotated sub-frame's content height runs along the original width
      // axis; for a table row the vertical extent is the original box width.
      return bw;
    }
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(isVert270 ? -Math.PI / 2 : Math.PI / 2);
    // After rotation the "width" direction of the new frame is the original height
    renderTextBody(ctx, { ...body, vert: 'horz' }, -bh / 2, -bw / 2, bh, bw, scale, shapeDefaultTextColor, 0, false, false, themeDefaultColor, slideNumber, rc, wrappedOnTextRun);
    ctx.restore();
    return;
  }
  const lPad = emuToPx(body.lIns, scale);
  const rPad = emuToPx(body.rIns, scale);
  const tPad = emuToPx(body.tIns, scale);
  const bPad = emuToPx(body.bIns, scale);
  // ECMA-376 §20.1.10.5 spAutoFit: "the shape's height and width should resize
  // to accommodate the text". The bbox is no longer a wrap boundary in that
  // mode — the shape grows to fit. Treating spAutoFit as wrap=none on the
  // horizontal axis keeps mixed-size sequences (e.g. "20代", "YoY+11.9%")
  // on one line, matching PowerPoint. Vertical growth is handled below by
  // the spAutoFit branch in the height calculation.
  // ECMA-376 §20.1.10.5 spAutoFit + §20.1.10.7 wrap interaction:
  // - `wrap="none"`: text never wraps (regardless of autoFit).
  // - `wrap="square"` + `<a:spAutoFit/>`: PowerPoint auto-fits the SHAPE to
  //   the text. If the text's natural single-line width fits the bbox, the
  //   shape stays at its current width and the text doesn't wrap (sample-2
  //   slide-13's "20代" textbox: ~50px text in ~70px bbox → no wrap). If
  //   the natural width exceeds the bbox, text wraps and the shape grows
  //   vertically (sample-2 slide-16's "1Q業績 要因" callouts: a single
  //   long Japanese paragraph that has to wrap within a fixed-width bbox).
  // - `wrap="square"` (default) without spAutoFit: always wrap.
  //
  // The pre-pass below measures the paragraphs' total natural width; if
  // every paragraph fits the bbox without wrapping, we keep the spAutoFit
  // "no-wrap" semantics. Otherwise we wrap normally so the long text
  // doesn't run off the side of the shape.
  const baseDoWrap = body.wrap !== 'none';
  const isSpAutoFit = body.autoFit === 'sp';
  const doWrap = isSpAutoFit
    ? (baseDoWrap && naturalWidthExceedsBbox(ctx, body, bw, lPad, rPad, scale, rc))
    : baseDoWrap;
  // ECMA-376 §20.1.10.34 numCol — distribute paragraphs across N text columns.
  const numCol = Math.max(1, body.numCol ?? 1);
  const spcColPx = emuToPx(body.spcCol ?? 0, scale);

  const bodyDefaultBold   = body.defaultBold   ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  const bodyDefaultColor = shapeDefaultTextColor ?? themeDefaultColor;

  // ── Pass 1: lay out all paragraphs ──────────────────────────────────────

  interface LineEntry {
    line: LayoutLine;
    linePx: number;       // spacing advancement (lineHeight + spaceAfter for last line)
    lineHeight: number;   // pure line height used for baseline positioning (without spaceAfter)
    topGapPx: number;     // spaceBefore for first line of paragraph
    textXOffset: number;  // additional X offset for first-line indent (non-bullet)
    bulletLabel: string;  // text to render as bullet ('' = none)
    bulletFont: string;
    bulletColor: string;
    bulletX: number;      // canvas X for bullet
    textX: number;        // canvas X for text
    textMaxW: number;     // max wrap width
    alignment: string;
    para: Paragraph;
  }

  // buildLayout runs Pass 1 at a given font scale (1.0 = normal; <1 = normAutoFit shrink)
  const buildLayout = (fontScale: number): { allLines: LineEntry[], totalHeight: number } => {
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale * fontScale;
  const allLines: LineEntry[] = [];
  let totalHeight = 0;

  // AutoNum counters per list level
  const autoNumCounters = new Map<number, number>();

  for (let paraIdx = 0; paraIdx < body.paragraphs.length; paraIdx++) {
    const para = body.paragraphs[paraIdx];
    const marLPx   = emuToPx(para.marL,   scale);
    const marRPx   = emuToPx(para.marR,   scale);
    const indentPx = emuToPx(para.indent, scale);

    // Para-level defaults (cascade: para defRPr → body default)
    const paraDefaultFontSizePx = para.defFontSize != null
      ? para.defFontSize * PT_TO_EMU * scale * fontScale : bodyDefaultFontSizePx;
    const paraDefaultColor = para.defColor
      ? hexToRgba(para.defColor) : bodyDefaultColor;

    // Bullet resolution
    const hasBullet = para.bullet.type === 'char' || para.bullet.type === 'autoNum';

    // Per ECMA-376 §21.1.2.4.13: when no buSz* is declared, the bullet takes
    // the first run's font size. Using paraDefaultFontSizePx here (the layout
    // lvl1pPr defRPr fallback, typically 18pt) oversizes the bullet so a
    // hanging indent calibrated against the run (12pt) can't contain it —
    // that's why the em-dash was overlapping the text.
    const firstRunSizePt = (() => {
      for (const r of para.runs) {
        if (r.type === 'text' && r.fontSize != null) return r.fontSize;
      }
      return null;
    })();
    const bulletBaseSizePx = firstRunSizePt != null
      ? firstRunSizePt * PT_TO_EMU * scale * fontScale
      : paraDefaultFontSizePx;

    // ECMA-376 §21.1.2.4.4 (CT_TextCharBullet) / §21.1.2.4.10 (buClrTx): when
    // no explicit `<a:buClr>` is present, the bullet inherits the *first run*'s
    // color, NOT the shape-level default text color. The two diverge on
    // templates where `<p:style><a:fontRef>` resolves to white (lt1) — runs
    // override that with their own `<a:rPr><a:solidFill>`, but bullets without
    // a buClr would otherwise pick up the white default and become invisible
    // (the slide-13 sample-2 regression).
    const firstRunColorHex = (() => {
      for (const r of para.runs) {
        if (r.type === 'text' && r.color) return r.color;
      }
      return null;
    })();
    const bulletInheritedColor = firstRunColorHex
      ? hexToRgba(firstRunColorHex)
      : paraDefaultColor;

    let bulletLabel  = '';
    let bulletFont   = buildFont(false, false, bulletBaseSizePx, 'sans-serif', rc);
    let bulletColor  = bulletInheritedColor;

    if (para.bullet.type === 'char') {
      const b = para.bullet;
      const bSizePx = b.sizePct != null
        ? bulletBaseSizePx * (b.sizePct / 100)
        : bulletBaseSizePx;
      bulletLabel = applySymbolFont(b.char, b.fontFamily ?? '');
      // If the char was mapped to a Unicode symbol, use sans-serif for reliable rendering.
      // Otherwise use the specified font (e.g. Wingdings on systems that have it).
      const convertedFamily = bulletLabel !== b.char ? 'sans-serif' : normalizeFontFamily(b.fontFamily ?? null, rc);
      bulletFont  = buildFont(false, false, bSizePx, convertedFamily, rc);
      bulletColor = b.color ? hexToRgba(b.color) : bulletInheritedColor;
      // Reset counters when switching to char bullets
      autoNumCounters.clear();
    } else if (para.bullet.type === 'autoNum') {
      const b = para.bullet;
      const lvl = para.lvl;
      if (!autoNumCounters.has(lvl)) {
        autoNumCounters.set(lvl, b.startAt ?? 1);
      } else {
        autoNumCounters.set(lvl, autoNumCounters.get(lvl)! + 1);
      }
      bulletLabel = formatAutoNum(autoNumCounters.get(lvl)!, b.numType);
      bulletFont  = buildFont(false, false, bulletBaseSizePx, 'sans-serif', rc);
      bulletColor = bulletInheritedColor;
    } else {
      // Not a list paragraph — reset autoNum counters
      autoNumCounters.clear();
    }

    // Text start X and wrap width.
    //
    // ECMA-376 §20.1.10.34 numCol — the bbox is split into N equal columns
    // separated by `spcCol` gutters, and paragraphs flow column-by-column.
    // For Pass 1 we lay out at the column width so wrapping & line counts are
    // correct; Pass 2 reuses the same lines and just shifts the X origin per
    // column. textX/bulletX below are relative to column 0; the column shift
    // is added later when we walk `allLines`.
    const colWidth = numCol > 1
      ? (bw - lPad - rPad - (numCol - 1) * spcColPx) / numCol
      : bw - lPad - rPad;
    const textX    = bx + lPad + marLPx;
    const bulletX  = bx + lPad + marLPx + indentPx;
    const textMaxW = colWidth - marLPx - marRPx;

    const maxW = doWrap ? textMaxW : Infinity;
    const lines = layoutParagraph(ctx, para, maxW, paraDefaultFontSizePx, paraDefaultColor, scale, marLPx, bodyDefaultBold, bodyDefaultItalic, fontScale, slideNumber, rc);

    // spaceBefore/After are in hundredths of a point → convert to canvas px
    const spaceBeforePx = para.spaceBefore != null ? (para.spaceBefore / 100) * PT_TO_EMU * scale * fontScale : 0;
    const spaceAfterPx  = para.spaceAfter  != null ? (para.spaceAfter  / 100) * PT_TO_EMU * scale * fontScale : 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const isFirst = i === 0;
      const isLast  = i === lines.length - 1;

      // Line height: use the max font size among rendered segments. The layout
      // default (paraDefaultFontSizePx) is used only as a fallback for empty
      // paragraphs — otherwise PowerPoint slide-layouts with placeholder
      // defaults like `defRPr sz="30000"` (300pt prompt-text marker) would
      // inflate lineHeight and push real 24pt runs far below the anchor.
      let maxSizePx = 0;
      for (const seg of line.segments) {
        // For an equation, the line must be at least as tall as its own font
        // size (so a short label like "y"/"p"/"z" gets the normal font-ascent
        // baseline and sits where PowerPoint puts it — not floated up by its
        // tight ink box), AND tall enough for its visual box (fractions, norms,
        // big operators) to fit inside the 1.2× leading.
        const effSize = seg.math
          ? Math.max(seg.sizePx, (seg.math.ascent + seg.math.descent) / 1.2)
          : seg.sizePx;
        if (effSize > maxSizePx) maxSizePx = effSize;
      }
      if (maxSizePx === 0) maxSizePx = paraDefaultFontSizePx;
      // Bullet font size also counts
      if (isFirst && bulletLabel) {
        ctx.font = bulletFont;
        const bm = ctx.measureText('M');
        const bSizeApprox = bm.actualBoundingBoxAscent + bm.actualBoundingBoxDescent;
        if (bSizeApprox > maxSizePx) maxSizePx = bSizeApprox;
      }

      let lineHeight: number;
      if (para.spaceLine) {
        if (para.spaceLine.type === 'pct') {
          // spcPct 100% = single line spacing = natural font leading ≈ 1.2× em
          lineHeight = maxSizePx * 1.2 * (para.spaceLine.val / 100000);
        } else {
          lineHeight = para.spaceLine.val * PT_TO_EMU * scale;
        }
      } else {
        lineHeight = maxSizePx * 1.2;
      }
      // normAutofit lnSpcReduction (ECMA-376 §21.1.2.1.3): PowerPoint reduces
      // each paragraph's line spacing by this fraction alongside the font
      // shrink. Apply it only when normAutofit actually stored a value.
      if (body.autoFit === 'norm' && body.lnSpcReduction != null) {
        lineHeight *= 1 - body.lnSpcReduction;
      }
      const linePx  = lineHeight + (isLast ? spaceAfterPx : 0);
      // ECMA-376 §21.1.2.2.6 (a:spcBef): paragraph "space before" is the gap
      // *between* paragraphs. PowerPoint suppresses it on the first paragraph
      // of a text body — otherwise placeholders whose layout-default `spcBef`
      // is 10 pt (sample-1 slide-5 "Figure 1." caption inherits this from the
      // layout body lstStyle) get pushed ~10 px below the placeholder top and
      // collide with the chart title sitting just below in the slide.
      const topGap  = isFirst && paraIdx > 0 ? spaceBeforePx : 0;
      // Non-bullet first-line indent
      const textXOffset = (!hasBullet && isFirst) ? indentPx : 0;

      allLines.push({
        line, linePx, lineHeight, topGapPx: topGap,
        textXOffset,
        bulletLabel: isFirst ? bulletLabel : '',
        bulletFont, bulletColor, bulletX,
        textX, textMaxW,
        alignment: para.alignment,
        para,
      });
      totalHeight += linePx + topGap;
    }
  }

  return { allLines, totalHeight };
  }; // end buildLayout

  let { allLines, totalHeight } = buildLayout(1.0);

  // ── normAutoFit ──────────────────────────────────────────────────────────
  // PowerPoint stores the font-shrink ratio it computed at edit time in
  // <a:normAutofit fontScale> (ECMA-376 §21.1.2.1.3). When present, apply it
  // directly — this reproduces PowerPoint's exact layout instead of guessing a
  // scale from our own (slightly different) text metrics. Only when no scale
  // was stored do we fall back to fitting the text by search.
  if (body.autoFit === 'norm') {
    if (body.fontScale != null && body.fontScale > 0) {
      if (body.fontScale < 1.0) ({ allLines, totalHeight } = buildLayout(body.fontScale));
    } else {
      const maxContentH = bh - tPad - bPad;
      if (totalHeight > maxContentH && maxContentH > 0) {
        let lo = 0.1, hi = 1.0;
        for (let i = 0; i < 6; i++) {
          const mid = (lo + hi) / 2;
          if (buildLayout(mid).totalHeight <= maxContentH) lo = mid; else hi = mid;
        }
        ({ allLines, totalHeight } = buildLayout(lo));
      }
    }
  }

  // ── measure-only: return the content height the text body needs ─────────
  // Used by renderTable to grow rows to fit their tallest cell (ECMA-376
  // §21.1.3.18: a:tr@h is a minimum). Returns padding + laid-out text height.
  if (measureOnly) {
    return tPad + totalHeight + bPad;
  }

  // ── anchor="b" with bh=0: auto-height growing upward from by ────────────
  // When cy=0 and anchor="b", off_y is the bottom anchor; shape grows upward.
  const anchor = body.verticalAnchor ?? 't';
  let effectiveBy = by;
  let effectiveBh: number;
  if (bh === 0 && anchor === 'b') {
    effectiveBh = tPad + totalHeight + bPad;
    effectiveBy = by - effectiveBh;
  } else {
    // ── Effective height (spAutoFit: shape expands to fit text) ─────────────
    const isSpAutoFit = body.autoFit === 'sp';
    effectiveBh = isSpAutoFit
      ? Math.max(bh, tPad + totalHeight + bPad)
      : bh;
  }

  // ── Vertical anchor ─────────────────────────────────────────────────────
  let cursorY: number;
  const contentH = Math.max(0, effectiveBh - tPad - bPad);
  if (anchor === 'ctr') {
    cursorY = effectiveBy + tPad + (contentH - totalHeight) / 2;
  } else if (anchor === 'b') {
    cursorY = effectiveBy + effectiveBh - bPad - totalHeight;
  } else {
    cursorY = effectiveBy + tPad;
  }

  // ── Pass 2: render ───────────────────────────────────────────────────────
  ctx.save();
  // penX / baseline are computed manually below, so the canvas text origin
  // must be normalized before fillText() or alignment/anchor math will drift.
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  // ECMA-376 §20.1.2.3.6 (a:bodyPr): PowerPoint does NOT clip text that
  // overflows its shape — the text simply renders past the shape bounds and
  // can overlap with adjacent elements. Our previous behavior clipped, which
  // cropped long text in fixed-height boxes. Only clip when the caller has
  // opted into wrap=none AND a finite x-axis rectangle (rare), which we
  // approximate here by skipping clipping entirely for the default bodyPr.
  // `body.wrap === "none"` means horizontal non-wrap; it doesn't affect
  // clipping per spec either, so we just don't clip.

  // Multi-column flow state. PowerPoint's actual behaviour with `numCol`:
  //
  // - If the total content height fits within a single column, ALL paragraphs
  //   stay in column 0 even when `numCol > 1`. (Sample-2 slide-13's "従来"
  //   box has 4 paragraphs in a `numCol="2"` body; PowerPoint stacks them
  //   vertically because they fit, not 2+2.)
  // - Otherwise, paragraphs are distributed *balanced* (each column gets
  //   roughly ⌈N/numCol⌉ entries) so parallel rows across columns line up.
  //   (Sample-2 slide-13's "新機能" box has 9 paragraphs that overflow a
  //   single column; the right column starts with a blank paragraph so the
  //   blue 利用履歴/価値観/属性詳細 sit at the same y as 年齢/性別/居住地.)
  //
  // A pure overflow-driven advance would either pack column 0 maximally and
  // shift the right column up (breaking row alignment) or never collapse
  // multi-col to single when content fits.
  //
  // We pre-compute, for each entry, which column it belongs to. Column 0
  // starts at the initial cursorY; each subsequent column resets the cursor
  // and shifts X by colWidth + spcCol.
  const colTopY = cursorY;
  const colWidthShift = numCol > 1
    ? (bw - lPad - rPad - (numCol - 1) * spcColPx) / numCol + spcColPx
    : 0;
  const colHeightCapacity = Math.max(0, effectiveBh - tPad - bPad);
  // When `bh === 0` (caller relies on auto-height) the capacity is unbounded
  // and content always fits, so a multi-col directive collapses to single col.
  const fitsInOneCol = bh === 0 || totalHeight <= colHeightCapacity + 0.5;
  const useMultiCol = numCol > 1 && !fitsInOneCol;
  const linesPerCol = useMultiCol ? Math.ceil(allLines.length / numCol) : allLines.length;
  let colIdx = 0;
  let entriesInCol = 0;

  for (const entry of allLines) {
    const { line, linePx, lineHeight, topGapPx, textXOffset, bulletLabel, bulletFont, bulletColor, alignment } = entry;
    // Balanced column advance: when the current column has reached its share
    // of paragraphs, jump to the next one. PowerPoint never breaks a single
    // line across columns and never spills past the last column — anything
    // beyond the last column simply runs past the bbox, matching PPT.
    if (numCol > 1 && colIdx < numCol - 1 && entriesInCol >= linesPerCol) {
      colIdx++;
      entriesInCol = 0;
      cursorY = colTopY;
    }
    cursorY += topGapPx;
    entriesInCol++;
    // §21.1.2.1.1 bodyPr@rtlCol: columns fill right-to-left — mirror the
    // visual column index. LTR bodies (rtlCol false/absent) are unchanged.
    const visCol = body.rtlCol ? numCol - 1 - colIdx : colIdx;
    const xShift = visCol * colWidthShift;
    const textX = entry.textX + xShift;
    const bulletX = entry.bulletX + xShift;
    const textMaxW = entry.textMaxW;

    // Bidi: base direction from a:pPr@rtl (the parser already flips the default
    // alignment l→r for rtl paragraphs). Engage only when the base is RTL or a
    // line carries strong-RTL characters, so LTR slides keep their exact path.
    const baseRtl = entry.para.rtl === true;
    const paraNeedsBidi = baseRtl || segmentsHaveRtl(line.segments);

    // Measure line for alignment AND baseline ascent in one pass.
    // actualBoundingBoxAscent gives the real font ascent for the rendered glyphs,
    // replacing the 0.8×lineHeight heuristic that over-estimates for CJK and
    // tall fonts, causing text to sit too low within the line box.
    let lineWidth = 0;
    let maxAscent = lineHeight * 0.8; // fallback when no segments
    for (const seg of line.segments) {
      if (seg.math) {
        lineWidth += seg.math.width;
        maxAscent = Math.max(maxAscent, seg.math.ascent);
        continue;
      }
      ctx.font = seg.font;
      const m = ctx.measureText(seg.text || 'M');
      const ls = seg.letterSpacingPx ?? 0;
      lineWidth += seg.text ? m.width + ls * seg.text.length : 0;
      if (m.actualBoundingBoxAscent > 0) {
        maxAscent = Math.max(maxAscent, m.actualBoundingBoxAscent);
      }
    }
    const baseline = cursorY + maxAscent;

    // Draw bullet. Under RTL the bullet hangs in the right gutter: mirror its
    // LTR offset (textX − bulletX, i.e. the bullet→text gap) about the text
    // column's right edge.
    if (bulletLabel) {
      ctx.font = bulletFont;
      ctx.fillStyle = bulletColor;
      if (paraNeedsBidi && baseRtl) {
        const prevDir = ctx.direction;
        ctx.direction = 'rtl';
        const bulletW = ctx.measureText(bulletLabel).width;
        ctx.fillText(bulletLabel, textX + textMaxW + (textX - bulletX) - bulletW, baseline);
        ctx.direction = prevDir;
      } else {
        ctx.fillText(bulletLabel, bulletX, baseline);
      }
    }

    const effectiveTextX = textX + textXOffset;
    let penX: number;
    if (alignment === 'ctr') {
      penX = effectiveTextX + (textMaxW - textXOffset - lineWidth) / 2;
    } else if (alignment === 'r') {
      penX = textX + textMaxW - lineWidth;
    } else {
      penX = effectiveTextX;
    }

    // Visual draw order: under bidi, reorder segments per UAX#9 (rule L2) and
    // draw each with ctx.direction matching its resolved direction. textAlign
    // is already 'left', so penX stays the segment's left edge.
    const visual: LineVisualOrder | null = paraNeedsBidi
      ? computeLineVisualOrder(line.segments, baseRtl)
      : null;
    const segCount = line.segments.length;
    for (let vi = 0; vi < segCount; vi++) {
      const li = visual ? visual.order[vi] : vi;
      const seg = line.segments[li];
      const segRtl = visual ? visual.rtl[li] : false;
      if (paraNeedsBidi) ctx.direction = segRtl ? 'rtl' : 'ltr';
      // ── Equation segment: draw the cached image instead of text ──────────
      if (seg.math) {
        const render = mathRenders.get(seg.math.nodes);
        const w = seg.math.width;
        const h = seg.math.ascent + seg.math.descent;
        if (render && w > 0 && h > 0) {
          const top = baseline - seg.math.ascent;
          const img = tintedMathImage(render, seg.color);
          ctx.drawImage(img, penX, top, w, h);
        }
        penX += w;
        continue;
      }
      ctx.font = seg.font;
      ctx.fillStyle = seg.color;
      // baseline shift: OOXML baseline in thousandths of a point; positive = superscript (up)
      const baselineShift = seg.baseline ? -(seg.baseline / 100000) * seg.sizePx : 0;
      const segBaseline = baseline + baselineShift;
      const ls = seg.letterSpacingPx ?? 0;

      // Run-level text shadow (rPr > effectLst > outerShdw). Set on the
      // context so the fillText below picks it up, then cleared after so
      // the outline / underline / strikethrough don't get shadowed too.
      // ECMA-376 §20.1.8.45 dir is degrees clockwise from east; the same
      // formula used by the shape-level shadow renderer above.
      const segShadow = seg.shadow;
      if (segShadow) {
        const dirRad = (segShadow.dir * Math.PI) / 180;
        const dist = emuToPx(segShadow.dist, scale);
        ctx.save();
        ctx.shadowColor = hexToRgba(segShadow.color, segShadow.alpha);
        ctx.shadowBlur = emuToPx(segShadow.blur, scale);
        ctx.shadowOffsetX = Math.cos(dirRad) * dist;
        ctx.shadowOffsetY = Math.sin(dirRad) * dist;
      }

      if (ls > 0 && seg.text.length > 1 && !segRtl) {
        // Draw glyph-by-glyph so each character advance is `measure + ls`.
        // Matches OOXML rPr @spc semantics — extra space added to each
        // character's advance, including after the last one.
        let cx = penX;
        for (const ch of seg.text) {
          ctx.fillText(ch, cx, segBaseline);
          cx += ctx.measureText(ch).width + ls;
        }
      } else if (ls > 0 && seg.text.length > 1) {
        // RTL segment with rPr @spc: per-glyph advance would break Arabic
        // cursive joining, so distribute the spacing via canvas letterSpacing
        // and draw the whole shaped segment in one fillText. The painted width
        // then matches the segW advance below (baseW + ls per character).
        const lctx = ctx as CanvasRenderingContext2D & { letterSpacing: string };
        try { lctx.letterSpacing = `${ls}px`; } catch { /* older engines */ }
        ctx.fillText(seg.text, penX, segBaseline);
        try { lctx.letterSpacing = '0px'; } catch { /* ignore */ }
      } else {
        ctx.fillText(seg.text, penX, segBaseline);
      }

      if (segShadow) ctx.restore();

      // Run-level text outline (rPr > a:ln). Strokes each glyph in addition
      // to the fill so the text reads as a thin lined character. ECMA-376
      // §20.1.2.2.24: `w` is EMU; convert to px via the same scale used for
      // fonts. Min 0.5 px so a 1-pt outline at small zoom levels stays
      // visible. Skip when width <= 0 (parser may emit width=0 for
      // explicitly empty <a:ln/>).
      const segOutline = seg.outline;
      if (segOutline && segOutline.width > 0) {
        ctx.save();
        ctx.lineWidth = Math.max(0.5, emuToPx(segOutline.width, scale));
        ctx.strokeStyle = segOutline.color ? `#${segOutline.color}` : seg.color;
        ctx.lineJoin = 'round';
        if (ls > 0 && seg.text.length > 1 && !segRtl) {
          let cx = penX;
          for (const ch of seg.text) {
            ctx.strokeText(ch, cx, segBaseline);
            cx += ctx.measureText(ch).width + ls;
          }
        } else if (ls > 0 && seg.text.length > 1) {
          const lctx = ctx as CanvasRenderingContext2D & { letterSpacing: string };
          try { lctx.letterSpacing = `${ls}px`; } catch { /* older engines */ }
          ctx.strokeText(seg.text, penX, segBaseline);
          try { lctx.letterSpacing = '0px'; } catch { /* ignore */ }
        } else {
          ctx.strokeText(seg.text, penX, segBaseline);
        }
        ctx.restore();
      }

      ctx.font = seg.font;
      const baseW = ctx.measureText(seg.text).width;
      const segW = baseW + (ls > 0 ? ls * seg.text.length : 0);

      if (onTextRun && seg.text) {
        onTextRun({
          text: seg.text,
          inShapeX: penX - bx,
          inShapeY: cursorY - by,
          w: segW,
          h: lineHeight,
          fontSize: seg.sizePx,
          font: seg.font,
          shapeX: bx,
          shapeY: by,
          shapeW: bw,
          shapeH: bh,
          rotation: shapeRotation,
        });
      }

      if (seg.underline) {
        drawUnderline(ctx, penX, segBaseline, segW, seg.sizePx, seg.underlineColor ?? seg.color, seg.underlineStyle);
      }

      if (seg.strikethrough) {
        const lineW = Math.max(1, seg.sizePx * 0.05);
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = lineW;
        ctx.setLineDash([]);
        if (seg.strikeDouble) {
          // Two parallel lines straddling the standard strike position,
          // separated by ~ 1.5× the line weight (visually distinct yet
          // staying inside the glyph's central band).
          const offset = lineW * 0.9;
          const yMid = segBaseline - seg.sizePx * 0.32;
          ctx.beginPath();
          ctx.moveTo(penX, yMid - offset);
          ctx.lineTo(penX + segW, yMid - offset);
          ctx.moveTo(penX, yMid + offset);
          ctx.lineTo(penX + segW, yMid + offset);
          ctx.stroke();
        } else {
          ctx.beginPath();
          ctx.moveTo(penX, segBaseline - seg.sizePx * 0.32);
          ctx.lineTo(penX + segW, segBaseline - seg.sizePx * 0.32);
          ctx.stroke();
        }
      }

      penX += segW;
    }
    if (paraNeedsBidi) ctx.direction = 'ltr'; // reset before tab-stop / next line

    // ── Tab-stop segments (right-aligned or centred at tab stop position) ──
    if (line.tabStop && line.tabStop.segments.length > 0) {
      const tabAbsX = bx + lPad + line.tabStop.px;
      let totalTabW = 0;
      for (const seg of line.tabStop.segments) {
        ctx.font = seg.font;
        const ls = seg.letterSpacingPx ?? 0;
        totalTabW += ctx.measureText(seg.text).width + ls * seg.text.length;
      }
      let tabPenX: number;
      if (line.tabStop.algn === 'r') {
        tabPenX = tabAbsX - totalTabW;
      } else if (line.tabStop.algn === 'ctr') {
        tabPenX = tabAbsX - totalTabW / 2;
      } else {
        tabPenX = tabAbsX;
      }
      for (const seg of line.tabStop.segments) {
        ctx.font = seg.font;
        ctx.fillStyle = seg.color;
        const tabLs = seg.letterSpacingPx ?? 0;
        if (tabLs > 0 && seg.text.length > 1) {
          let cx = tabPenX;
          for (const ch of seg.text) {
            ctx.fillText(ch, cx, baseline);
            cx += ctx.measureText(ch).width + tabLs;
          }
        } else {
          ctx.fillText(seg.text, tabPenX, baseline);
        }
        ctx.font = seg.font;
        const tabSegW = ctx.measureText(seg.text).width + tabLs * seg.text.length;
        if (onTextRun && seg.text) {
          onTextRun({
            text: seg.text,
            inShapeX: tabPenX - bx,
            inShapeY: cursorY - by,
            w: tabSegW,
            h: lineHeight,
            fontSize: seg.sizePx,
            font: seg.font,
            shapeX: bx,
            shapeY: by,
            shapeW: bw,
            shapeH: bh,
            rotation: shapeRotation,
          });
        }
        tabPenX += tabSegW;
      }
    }

    cursorY += linePx;
  }

  ctx.restore();
}

// Decoded-image cache keyed by data URL. Decoding an inlined base64 image to an
// ImageBitmap is expensive, and the same picture is re-decoded on every render
// (each scroll / resize / interaction). Cache the decode — the Promise, so
// concurrent first-renders dedupe — and reuse it. Bounded FIFO so a long
// session or many decks can't grow without limit.
const IMAGE_BITMAP_CACHE_MAX = 256;
const imageBitmapCache = new Map<string, Promise<ImageBitmap>>();

function getCachedBitmap(dataUrl: string): Promise<ImageBitmap> {
  const existing = imageBitmapCache.get(dataUrl);
  if (existing) {
    // Refresh LRU position.
    imageBitmapCache.delete(dataUrl);
    imageBitmapCache.set(dataUrl, existing);
    return existing;
  }
  const p = fetch(dataUrl).then((r) => r.blob()).then((b) => createImageBitmap(b));
  // Don't poison the cache on a transient decode failure.
  p.catch(() => imageBitmapCache.delete(dataUrl));
  imageBitmapCache.set(dataUrl, p);
  if (imageBitmapCache.size > IMAGE_BITMAP_CACHE_MAX) {
    const oldestKey = imageBitmapCache.keys().next().value as string;
    const oldest = imageBitmapCache.get(oldestKey);
    imageBitmapCache.delete(oldestKey);
    oldest?.then((b) => b.close()).catch(() => {});
  }
  return p;
}

async function renderPicture(
  ctx: CanvasRenderingContext2D,
  el: PictureElement,
  scale: number
) {
  try {
    const bitmap = await getCachedBitmap(el.dataUrl);
    ctx.save();
    if (el.alpha != null) ctx.globalAlpha *= el.alpha;
    const x = emuToPx(el.x, scale);
    const y = emuToPx(el.y, scale);
    const w = emuToPx(el.width, scale);
    const h = emuToPx(el.height, scale);
    if (el.rotation !== 0 || el.flipH || el.flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate((el.rotation * Math.PI) / 180);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }

    // ── srcRect sub-rectangle (ECMA-376 a:srcRect). Edge values are fractions
    // of source dims (negative values mean extend past the image, duplicating
    // edge pixels in OOXML; we clamp to [0,1]). Resolve once so both the live
    // paint and the effect aux paints share identical crop coordinates.
    const sr = el.srcRect;
    let crop: { sx: number; sy: number; sw: number; sh: number } | null = null;
    if (sr && (sr.l || sr.t || sr.r || sr.b)) {
      const bw = bitmap.width, bh = bitmap.height;
      const sl = Math.max(0, Math.min(1, sr.l ?? 0));
      const st = Math.max(0, Math.min(1, sr.t ?? 0));
      const srR = Math.max(0, Math.min(1, sr.r ?? 0));
      const sbB = Math.max(0, Math.min(1, sr.b ?? 0));
      const sx = sl * bw;
      const sy = st * bh;
      crop = {
        sx,
        sy,
        sw: Math.max(1, bw - sx - srR * bw),
        sh: Math.max(1, bh - sy - sbB * bh),
      };
    }

    // Apply the picture clip (roundRect / custGeom) to an arbitrary target.
    // ECMA-376 §20.1.9.8: a `<p:pic>` may carry `<a:custGeom>` defining a
    // non-rectangular silhouette the bitmap is trimmed to (e.g. a laptop frame).
    const applyClip = (target: CanvasRenderingContext2D): void => {
      if (el.clipAdjust != null) {
        const minDim = Math.min(w, h);
        const r = (el.clipAdjust / 100000) * minDim;
        target.beginPath();
        target.roundRect(x, y, w, h, r);
        target.clip();
      } else if (el.custGeom && el.custGeom.length > 0) {
        buildCustomPath(target, el.custGeom, x, y, w, h);
        target.clip();
      }
    };

    // Draw the (clipped, optionally cropped) bitmap into a target context. This
    // is the picture "body" that the effect helpers re-paint onto aux canvases,
    // so reflections/soft edges mirror the real image rather than a flat shape.
    const paintImage = (target: CanvasRenderingContext2D): void => {
      target.save();
      applyClip(target);
      if (crop) {
        target.drawImage(bitmap, crop.sx, crop.sy, crop.sw, crop.sh, x, y, w, h);
      } else {
        target.drawImage(bitmap, x, y, w, h);
      }
      target.restore();
    };

    // Flat opaque silhouette of the clipped picture rectangle, used as the mask
    // for innerShdw / softEdge. Falls back to the bounding rect when the picture
    // is a plain rectangle (no clip path).
    const paintMask = (target: CanvasRenderingContext2D, color: string): void => {
      target.save();
      applyClip(target);
      target.fillStyle = color;
      target.fillRect(x, y, w, h);
      target.restore();
    };

    // ── effectLst (§19.3.1.37 routes p:pic's spPr through CT_ShapeProperties,
    // so §20.1.8.16 effects apply to images). Same sequence as the p:sp path.
    const deviceW = (ctx.canvas as { width: number }).width || 0;
    const deviceH = (ctx.canvas as { height: number }).height || 0;
    const liveTransform = ctx.getTransform();
    const det = Math.abs(
      liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c,
    );
    const devScale = det > 0 ? Math.sqrt(det) : 1;
    const effBBox = { x: x * devScale, y: y * devScale, w: w * devScale, h: h * devScale };
    const effScale = scale * devScale; // EMU → device px
    const applyLiveTransform = (c: CanvasRenderingContext2D) => c.setTransform(liveTransform);
    const haveAux = deviceW > 0 && deviceH > 0;

    // Reflection sits below the picture — paint it first. §20.1.8.50. The aux
    // paint bakes in the live rotation/flip via setTransform, so the blit runs
    // at identity.
    if (el.reflection && haveAux) {
      ctx.save();
      ctx.setTransform(new DOMMatrix());
      applyReflection(
        ctx,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintImage(c as CanvasRenderingContext2D); },
        effBBox, el.reflection, effScale, deviceW, deviceH,
      );
      ctx.restore();
    }

    // outerShdw / glow use the single Canvas shadow slot, cast by the image's
    // own opaque pixels. Outer shadow wins when both are present (as in p:sp).
    if (el.shadow) applyShadow(ctx, el.shadow, scale);
    else if (el.glow) applyGlow(ctx, el.glow, scale);

    // softEdge feathers the whole picture, REPLACING the direct body paint
    // (§20.1.8.53). The shadow/glow set above is carried into the aux paint via
    // setTransform of the same live context, so it still casts.
    if (el.softEdge && haveAux) {
      ctx.save();
      ctx.setTransform(new DOMMatrix());
      applySoftEdge(
        ctx,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintImage(c as CanvasRenderingContext2D); },
        effBBox, el.softEdge, effScale, deviceW, deviceH,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintMask(c as CanvasRenderingContext2D, '#000'); },
      );
      ctx.restore();
    } else {
      paintImage(ctx);
    }
    if (el.shadow || el.glow) clearShadow(ctx);

    // innerShdw casts inward, on top of the picture (§20.1.8.40).
    if (el.innerShadow && haveAux) {
      ctx.save();
      ctx.setTransform(new DOMMatrix());
      applyInnerShadow(
        ctx,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintMask(c as CanvasRenderingContext2D, '#000'); },
        effBBox, el.innerShadow, effScale, deviceW, deviceH,
      );
      ctx.restore();
    }

    ctx.restore();
    // bitmap is owned by getCachedBitmap's cache — do not close it here.
  } catch {
    // silently skip broken images
  }
}

async function renderMedia(
  ctx: CanvasRenderingContext2D,
  el: MediaElement,
  scale: number,
  fetchMedia?: (path: string) => Promise<Blob>,
  skipControls?: boolean,
) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  let drewPoster = false;
  if (el.posterPath && fetchMedia) {
    try {
      const blob = await fetchMedia(el.posterPath);
      const typed = el.posterMimeType ? new Blob([blob], { type: el.posterMimeType }) : blob;
      const bitmap = await createImageBitmap(typed);
      ctx.drawImage(bitmap, x, y, w, h);
      bitmap.close();
      drewPoster = true;
    } catch {
      // fall through to plain fill
    }
  }
  if (!drewPoster) {
    ctx.fillStyle = el.mediaKind === 'video' ? '#111' : '#f0f0f0';
    ctx.fillRect(x, y, w, h);
  }

  if (skipControls) return;

  drawPlayBadge(ctx, x + w / 2, y + h / 2, w, h, 'paused');
}

// ===== Table renderer =====

/**
 * Re-stroke a straight connector under one of the ECMA-376 §20.1.8.42
 * compound-line styles. Each sub-line gets its own offset along the
 * perpendicular and its own thickness; widths are taken from Office's
 * usual interpretation:
 *
 *   dbl       → two equal-thickness lines, each w/3, gap w/3
 *   thinThick → thin (w/4) + gap (w/4) + thick (w/2)
 *   thickThin → thick (w/2) + gap (w/4) + thin (w/4)
 *   tri       → three lines: thin(w/5), thick(w/5*3 ≈ 3w/5), thin(w/5)
 *
 * The starting position of each sub-line is computed so the *outer envelope*
 * stays w wide — i.e. centred on the original geometry. The base stroke
 * already drew at the centre line; we erase it (destination-out) and then
 * paint the compound sub-lines on top so dash/headEnd remain consistent.
 */
function drawCompoundLine(
  ctx: CanvasRenderingContext2D,
  start: { x: number; y: number },
  end: { x: number; y: number },
  stroke: Stroke,
  cmpd: string,
  scale: number,
): void {
  const totalW = Math.max(0.5, emuToPx(stroke.width, scale));
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const len = Math.hypot(dx, dy);
  if (len === 0) return;
  const px = -dy / len;
  const py = dx / len;

  // sub: relative position along the perpendicular axis (-1..1, with 0 being
  // the centre line) and width as a fraction of totalW.
  type Sub = { offset: number; widthFrac: number };
  let subs: Sub[];
  switch (cmpd) {
    case 'dbl':
      subs = [
        { offset: -1 / 3, widthFrac: 1 / 3 },
        { offset:  1 / 3, widthFrac: 1 / 3 },
      ];
      break;
    case 'thinThick':
      subs = [
        { offset: -3 / 8, widthFrac: 1 / 4 },
        { offset:  1 / 4, widthFrac: 1 / 2 },
      ];
      break;
    case 'thickThin':
      subs = [
        { offset: -1 / 4, widthFrac: 1 / 2 },
        { offset:  3 / 8, widthFrac: 1 / 4 },
      ];
      break;
    case 'tri':
      subs = [
        { offset: -2 / 5, widthFrac: 1 / 5 },
        { offset:  0,     widthFrac: 3 / 5 },
        { offset:  2 / 5, widthFrac: 1 / 5 },
      ];
      break;
    default:
      return;
  }

  ctx.save();
  // 1. Erase the centre line that was already drawn at full width.
  ctx.globalCompositeOperation = 'destination-out';
  ctx.strokeStyle = '#000';
  ctx.lineWidth = totalW + 0.5; // small overshoot to fully erase antialiasing fringe
  ctx.setLineDash([]);
  ctx.beginPath();
  ctx.moveTo(start.x, start.y); ctx.lineTo(end.x, end.y);
  ctx.stroke();

  // 2. Paint each sub-line.
  ctx.globalCompositeOperation = 'source-over';
  ctx.strokeStyle = hexToRgba(stroke.color);
  for (const sub of subs) {
    const ox = px * (totalW * sub.offset);
    const oy = py * (totalW * sub.offset);
    ctx.lineWidth = Math.max(0.5, totalW * sub.widthFrac);
    ctx.beginPath();
    ctx.moveTo(start.x + ox, start.y + oy);
    ctx.lineTo(end.x + ox, end.y + oy);
    ctx.stroke();
  }
  ctx.restore();
}

/** Draw an arrowhead at `tip` pointing in `angle` radians (0 = right). */
function drawArrowHead(
  ctx: CanvasRenderingContext2D,
  tipX: number,
  tipY: number,
  angle: number,
  arrowEnd: { type: string; w: string; len: string },
  stroke: Stroke,
  scale: number,
): void {
  if (arrowEnd.type === 'none') return;
  const lw = Math.max(0.5, emuToPx(stroke.width, scale));
  // Arrowhead size is implementation-defined: ECMA-376 §20.1.10.32/.33 only name
  // the w/len steps (sm/med/lg) as "relative", not exact ratios. These multiples
  // of line width are calibrated to PowerPoint's rendering (verified against the
  // sample-9 SmartArt timeline PDF, a lg/lg triangle on a 1pt line ≈ 8× wide).
  const wMul = arrowEnd.w   === 'sm' ? 4 : arrowEnd.w   === 'lg' ? 8 : 6;
  const lMul = arrowEnd.len === 'sm' ? 4 : arrowEnd.len === 'lg' ? 8 : 6;
  const halfW = lw * wMul / 2;
  const len   = lw * lMul;
  const color = hexToRgba(stroke.color);

  ctx.save();
  ctx.translate(tipX, tipY);
  ctx.rotate(angle);
  ctx.fillStyle   = color;
  ctx.strokeStyle = color;
  ctx.lineWidth   = lw;
  ctx.setLineDash([]);
  ctx.beginPath();
  switch (arrowEnd.type) {
    case 'triangle':
    case 'stealth':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len, -halfW);
      ctx.lineTo(-len,  halfW);
      ctx.closePath();
      ctx.fill();
      break;
    case 'arrow':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len, -halfW);
      ctx.moveTo(0, 0);
      ctx.lineTo(-len,  halfW);
      ctx.stroke();
      break;
    case 'diamond':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len / 2, -halfW);
      ctx.lineTo(-len, 0);
      ctx.lineTo(-len / 2,  halfW);
      ctx.closePath();
      ctx.fill();
      break;
    case 'oval':
      ctx.ellipse(-len / 2, 0, len / 2, halfW, 0, 0, Math.PI * 2);
      ctx.fill();
      break;
  }
  ctx.restore();
}

function applyStroke(ctx: CanvasRenderingContext2D, stroke: Stroke | null, scale: number) {
  // `scale` is EMU → px factor (canvasWidthPx / slideWidthEMU).
  applyStrokeCore(ctx, stroke, scale);
}

// ─── Chart rendering ────────────────────────────────────────────────────────
// Chart rendering is delegated to @silurus/ooxml-core's unified renderer.
// See renderChart(ctx, el, scale) below.


// ─── Table rendering ─────────────────────────────────────────────────────────

function renderTable(ctx: CanvasRenderingContext2D, el: TableElement, scale: number, slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null }) {
  const x0 = emuToPx(el.x, scale);
  const y0 = emuToPx(el.y, scale);

  // Convert col widths to pixels.
  const colWidths = el.cols.map(c => emuToPx(c, scale));
  const numCols = colWidths.length;

  // Spanned width starting at column `ci` for `span` columns.
  const spannedWidth = (ci: number, span: number): number => {
    let w = 0;
    for (let s = 0; s < span; s++) w += colWidths[ci + s] ?? 0;
    return w;
  };

  // ── Row heights: ECMA-376 §21.1.3.18 (a:tr@h) is a MINIMUM ────────────────
  // PowerPoint grows a row to fit its tallest cell's laid-out text (like
  // Word's "at least" line rule). A literal h=0 therefore becomes
  // content-driven. We measure each cell's text body at its spanned width
  // (reusing the same renderTextBody machinery via measureOnly) and take
  // max(tr@h, tallest single-row cell content). A rowSpan cell distributes
  // its content height across the rows it covers so it doesn't inflate the
  // first row.
  const rowHeights = el.rows.map(r => emuToPx(r.height, scale));

  // First pass: single-row (rowSpan ≤ 1) cells set their own row's minimum.
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.hMerge || cell.vMerge) continue;
      if ((cell.rowSpan || 1) > 1) continue;
      if (!cell.textBody) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      const needed = (renderTextBody(
        ctx, cell.textBody, 0, 0, cellW, 0, scale, null, 0, false, false,
        '#000000', slideNumber, rc, undefined, true,
      ) as number) || 0;
      if (needed > rowHeights[ri]) rowHeights[ri] = needed;
    }
  }

  // Second pass: rowSpan cells. If the content needs more than the sum of the
  // rows it spans, distribute the deficit across those rows so the merged area
  // grows without inflating any single row beyond what its own content needs.
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.hMerge || cell.vMerge) continue;
      const span = cell.rowSpan || 1;
      if (span <= 1 || !cell.textBody) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      const needed = (renderTextBody(
        ctx, cell.textBody, 0, 0, cellW, 0, scale, null, 0, false, false,
        '#000000', slideNumber, rc, undefined, true,
      ) as number) || 0;
      let have = 0;
      for (let s = 0; s < span && ri + s < rowHeights.length; s++) have += rowHeights[ri + s];
      if (needed > have) {
        const extra = (needed - have) / span;
        for (let s = 0; s < span && ri + s < rowHeights.length; s++) rowHeights[ri + s] += extra;
      }
    }
  }

  // ── Column x-positions ────────────────────────────────────────────────────
  // ECMA-376 §21.1.3.13 (a:tblPr@rtl): a right-to-left table places column 0
  // at the right edge, columns advancing leftward. We precompute the left
  // pixel edge of each column so RTL is a coordinate flip (no ctx.scale(-1,1)
  // which would mirror text and borders).
  const tableW = colWidths.reduce((a, b) => a + b, 0);
  const colLeft: number[] = new Array(numCols);
  if (el.rtl) {
    let right = x0 + tableW;
    for (let ci = 0; ci < numCols; ci++) {
      right -= colWidths[ci];
      colLeft[ci] = right;
    }
  } else {
    let left = x0;
    for (let ci = 0; ci < numCols; ci++) {
      colLeft[ci] = left;
      left += colWidths[ci];
    }
  }

  // Left pixel edge of a cell spanning columns [ci, ci+span). Under RTL the
  // span grows leftward, so the merged cell's left edge is the leftmost column.
  const spannedLeft = (ci: number, span: number): number =>
    el.rtl ? colLeft[ci + span - 1] : colLeft[ci];

  const rowTop: number[] = new Array(el.rows.length);
  {
    let y = y0;
    for (let ri = 0; ri < el.rows.length; ri++) { rowTop[ri] = y; y += rowHeights[ri]; }
  }

  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    const rowY = rowTop[ri];

    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];

      // Merged cells that are continuations: skip drawing
      if (cell.hMerge || cell.vMerge) {
        continue;
      }

      // Cell size: span multiple columns/rows
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      let cellH = 0;
      for (let span = 0; span < (cell.rowSpan || 1); span++) {
        cellH += rowHeights[ri + span] ?? 0;
      }
      const colX = spannedLeft(ci, cell.gridSpan || 1);

      // Fill
      const fillColor = resolveFill(cell.fill);
      if (fillColor) {
        ctx.fillStyle = fillColor;
        ctx.fillRect(colX, rowY, cellW, cellH);
      }

      // Text body — default run colour comes from the table style's tcTxStyle
      // (e.g. white header text on an accent fill); a run's explicit colour wins.
      if (cell.textBody) {
        const cellDefaultColor = cell.textColor ? hexToRgba(cell.textColor) : null;
        renderTextBody(ctx, cell.textBody, colX, rowY, cellW, cellH, scale, cellDefaultColor, 0, false, false, '#000000', slideNumber, rc);
      }

      // Borders
      ctx.save();
      if (cell.borderT) {
        applyStroke(ctx, cell.borderT, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY);
        ctx.lineTo(colX + cellW, rowY);
        ctx.stroke();
      }
      if (cell.borderB) {
        applyStroke(ctx, cell.borderB, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY + cellH);
        ctx.lineTo(colX + cellW, rowY + cellH);
        ctx.stroke();
      }
      if (cell.borderL) {
        applyStroke(ctx, cell.borderL, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY);
        ctx.lineTo(colX, rowY + cellH);
        ctx.stroke();
      }
      if (cell.borderR) {
        applyStroke(ctx, cell.borderR, scale);
        ctx.beginPath();
        ctx.moveTo(colX + cellW, rowY);
        ctx.lineTo(colX + cellW, rowY + cellH);
        ctx.stroke();
      }
      // Diagonal borders: top-left→bottom-right and bottom-left→top-right
      if (cell.diagonalTL) {
        applyStroke(ctx, cell.diagonalTL, scale);
        ctx.beginPath();
        ctx.moveTo(colX, rowY);
        ctx.lineTo(colX + cellW, rowY + cellH);
        ctx.stroke();
      }
      if (cell.diagonalTR) {
        applyStroke(ctx, cell.diagonalTR, scale);
        ctx.beginPath();
        ctx.moveTo(colX + cellW, rowY);
        ctx.lineTo(colX, rowY + cellH);
        ctx.stroke();
      }
      ctx.restore();
    }
  }
}

// ===== Public API =====

export type { RenderOptions } from './types';

/**
 * Internal render options: the shared {@link RenderOptions} plus the opt-in
 * `math` engine. `math` is internal plumbing — the headless {@link
 * PptxPresentation} injects it once at load and threads it here on each draw,
 * so the public `RenderSlideOptions` deliberately does not expose it.
 */
type SlideRenderOptions = RenderOptions & { math?: MathRenderer };

/**
 * Render a single slide onto a <canvas> element.
 * Returns the canvas for convenience.
 */
export async function renderSlide(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: SlideRenderOptions = {},
  onTextRun?: TextRunCallback,
): Promise<HTMLCanvasElement | OffscreenCanvas> {
  // Cancellation guard. renderSlide is async (it awaits image / equation decode),
  // so rapid navigation can start a newer render of the SAME canvas before this
  // one finishes. Both would `canvas.width = …` (clear) and then draw, and their
  // draws interleave at the await points — ghosting multiple slides together.
  // Stamp a per-canvas token; once a newer render supersedes us, stop drawing at
  // the next await so only the latest render's output survives.
  const tokenHost = canvas as unknown as { __pptxRenderToken?: number };
  const myToken = (tokenHost.__pptxRenderToken = (tokenHost.__pptxRenderToken ?? 0) + 1);
  const superseded = () => tokenHost.__pptxRenderToken !== myToken;

  const targetWidth = opts.width ?? ((canvas instanceof HTMLCanvasElement ? canvas.offsetWidth : 0) || 960);
  const scale = targetWidth / slideWidth;
  const canvasW = Math.round(targetWidth);
  const canvasH = Math.round(slideHeight * scale);

  const dpr = opts.dpr ?? (typeof window !== 'undefined' ? (window.devicePixelRatio || 1) : 1);
  canvas.width  = canvasW * dpr;
  canvas.height = canvasH * dpr;
  // CSS size only applies to the visible HTMLCanvasElement (not OffscreenCanvas)
  if (typeof HTMLCanvasElement !== 'undefined' && canvas instanceof HTMLCanvasElement) {
    canvas.style.width = `${canvasW}px`;
    // Mirror the docx renderer: when callers use `renderSlide(canvas, ...)`
    // directly without the {@link PptxViewer} wrapper, set `display:block`
    // as a safety net so the inline-element baseline does not leave a
    // descender gap below the canvas. Respect any user-specified value.
    if (!canvas.style.display) canvas.style.display = 'block';
  }

  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) throw new Error('Could not get 2D context');
  ctx.scale(dpr, dpr);

  const rc: RenderContext = {
    themeMajorFont: opts.majorFont ?? null,
    themeMinorFont: opts.minorFont ?? null,
    themeHlinkColor: opts.hlinkColor ?? null,
  };

  renderBackground(ctx, slide.background, canvasW, canvasH);

  // Pre-rasterize any equations so the synchronous text layout can place them.
  // `math` is the engine injected once at PptxPresentation.load and threaded in
  // here; without it, equations are skipped and the asset never enters the bundle.
  if (opts.math) await prepareSlideMath(slide, opts.math);
  if (superseded()) return canvas;

  const themeDefaultColor = opts.defaultTextColor
    ? `#${opts.defaultTextColor}`
    : '#000000';

  const slideNumber = slide.slideNumber;
  for (const el of slide.elements) {
    // A newer render of this canvas started while we awaited an image/equation —
    // stop so we don't paint this (now stale) slide over the newer one.
    if (superseded()) return canvas;
    if (el.type === 'shape') {
      renderShape(ctx, el, scale, themeDefaultColor, slideNumber, rc, onTextRun);
    } else if (el.type === 'picture') {
      await renderPicture(ctx, el, scale);
    } else if (el.type === 'table') {
      renderTable(ctx, el, scale, slideNumber, rc);
    } else if (el.type === 'media') {
      await renderMedia(ctx, el, scale, opts.fetchMedia, opts.skipMediaControls);
    } else if (el.type === 'chart') {
      // OOXML: 1pt = 12700 EMU. The slide renderer's `scale` is px-per-EMU,
      // so PT_TO_EMU * scale gives pixels-per-point at the current display size.
      const chartPtToPx = PT_TO_EMU * scale;
      renderChart(
        ctx,
        {
          chartType: el.chartType,
          title: el.title,
          categories: el.categories,
          series: el.series,
          showDataLabels: el.showDataLabels,
          valMin: el.valMin,
          valMax: el.valMax,
          catAxisTitle: null,
          valAxisTitle: null,
          catAxisHidden: el.catAxisHidden,
          valAxisHidden: el.valAxisHidden,
          catAxisLineHidden: el.catAxisLineHidden ?? false,
          valAxisLineHidden: el.valAxisLineHidden ?? false,
          catAxisFontColor: el.catAxisFontColor ?? null,
          valAxisFontColor: el.valAxisFontColor ?? null,
          catAxisLineColor: el.catAxisLineColor ?? null,
          catAxisLineWidthEmu: el.catAxisLineWidthEmu ?? null,
          valAxisLineColor: el.valAxisLineColor ?? null,
          valAxisLineWidthEmu: el.valAxisLineWidthEmu ?? null,
          plotAreaBg: el.plotAreaBg,
          chartBg: el.chartBg,
          showLegend: el.showLegend,
          legendPos: el.legendPos ?? null,
          catAxisCrossBetween: el.catAxisCrossBetween,
          valAxisMajorTickMark: el.valAxisMajorTickMark,
          catAxisMajorTickMark: el.catAxisMajorTickMark,
          titleFontSizeHpt: el.titleFontSizeHpt,
          titleFontColor: el.titleFontColor ?? null,
          titleFontFace: el.titleFontFace ?? null,
          catAxisFontSizeHpt: el.catAxisFontSizeHpt,
          valAxisFontSizeHpt: el.valAxisFontSizeHpt,
          dataLabelFontSizeHpt: el.dataLabelFontSizeHpt,
          subtotalIndices: el.subtotalIndices,
          barGapWidth: el.barGapWidth ?? null,
          barOverlap: el.barOverlap ?? null,
          dataLabelPosition: el.dataLabelPosition ?? null,
          dataLabelFontColor: el.dataLabelFontColor ?? null,
          dataLabelFormatCode: el.dataLabelFormatCode ?? null,
          valAxisFormatCode: el.valAxisFormatCode ?? null,
          plotAreaManualLayout: el.plotAreaManualLayout ?? null,
          scatterStyle: el.scatterStyle ?? null,
          radarStyle: el.radarStyle ?? null,
        },
        {
          x: emuToPx(el.x, scale),
          y: emuToPx(el.y, scale),
          w: emuToPx(el.width, scale),
          h: emuToPx(el.height, scale),
        },
        chartPtToPx,
      );
    }
  }

  return canvas;
}
