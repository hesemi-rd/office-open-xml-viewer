import type { MathConstants } from './math-table';
import type { MathFont } from './font';
import type { MathNode, MathStyle } from '../types/math';

export type MathLevel = 'display' | 'text' | 'script' | 'scriptScript';

/** Per-glyph ink metrics in CSS px (from a canvas `measureText`, when available). */
export type MeasureGlyph = (
  text: string,
  sizePx: number,
  style: MathStyle,
) => { width: number; ascent: number; descent: number };

export interface MathLayoutCtx {
  font: MathFont;
  consts: MathConstants;
  /** em size in CSS px at this level. */
  fontSizePx: number;
  level: MathLevel;
  /** Optional real per-glyph metrics (canvas-backed). Falls back to font globals. */
  measureGlyph?: MeasureGlyph;
}

export type DrawOp =
  | { type: 'glyph'; text: string; style: MathStyle; x: number; y: number; sizePx: number }
  | { type: 'rule'; x: number; y: number; w: number; h: number }
  /** Polyline stroke (e.g. a synthesized radical sign). Points are relative to the
   *  box origin; y grows down with the baseline at 0. */
  | { type: 'stroke'; points: { x: number; y: number }[]; lineWidth: number };

export interface MathBox {
  width: number;
  /** distance above baseline (px, positive). */
  ascent: number;
  /** distance below baseline (px, positive). */
  descent: number;
  /** ops positioned relative to this box's origin: x grows right, y grows down,
   *  baseline at y = 0. Content above the baseline has negative y. */
  ops: DrawOp[];
}

// ── Atom classes (TeX math) ─────────────────────────────────────────────────
export type AtomClass = 'ord' | 'op' | 'bin' | 'rel' | 'open' | 'close' | 'punct' | 'inner';

interface Atom {
  box: MathBox;
  cls: AtomClass;
}

const BIN_CHARS = '+−±∓×÷·∗⋅∘∙';
const REL_CHARS = '=≠<>≤≥≈≡∼≅≃→←↔⇒∈∉⊂⊆⊃⊇∝≪≫⊥';
const OPEN_CHARS = '([{⟨⌈⌊';
const CLOSE_CHARS = ')]}⟩⌉⌋';
const PUNCT_CHARS = ',;';
const BIGOP_CHARS = '∑∏∐∫∮⋃⋂⨀';

function classifyChar(ch: string): AtomClass {
  if (BIN_CHARS.includes(ch)) return 'bin';
  if (REL_CHARS.includes(ch)) return 'rel';
  if (OPEN_CHARS.includes(ch)) return 'open';
  if (CLOSE_CHARS.includes(ch)) return 'close';
  if (PUNCT_CHARS.includes(ch)) return 'punct';
  if (BIGOP_CHARS.includes(ch)) return 'op';
  return 'ord';
}

// TeX inter-atom spacing (math units, 1mu = em/18). Text/display style. Missing = 0.
const SPACE: Record<AtomClass, Partial<Record<AtomClass, number>>> = {
  ord: { op: 3, bin: 4, rel: 5, inner: 3 },
  op: { ord: 3, op: 3, rel: 5, inner: 3 },
  bin: { ord: 4, op: 4, open: 4, inner: 4 },
  rel: { ord: 5, op: 5, open: 5, inner: 5 },
  open: {},
  close: { op: 3, bin: 4, rel: 5, inner: 3 },
  punct: { ord: 3, op: 3, rel: 3, open: 3, close: 3, punct: 3, inner: 3 },
  inner: { ord: 3, op: 3, bin: 4, rel: 5, open: 3, punct: 3, inner: 3 },
};

const muToPx = (mu: number, ctx: MathLayoutCtx) => (mu / 18) * ctx.fontSizePx;

/** font units -> px at the current level. */
const fu = (ctx: MathLayoutCtx, v: number) => (v * ctx.fontSizePx) / ctx.font.unitsPerEm;

function scaleFor(level: MathLevel, c: MathConstants): number {
  if (level === 'script') return c.scriptPercentScaleDown / 100;
  if (level === 'scriptScript') {
    return (c.scriptPercentScaleDown / 100) * (c.scriptScriptPercentScaleDown / 100);
  }
  return 1;
}

function childLevel(level: MathLevel): MathLevel {
  return level === 'display' || level === 'text' ? 'script' : 'scriptScript';
}

function scriptCtx(ctx: MathLayoutCtx): MathLayoutCtx {
  const level = childLevel(ctx.level);
  const baseEm =
    ctx.level === 'display' || ctx.level === 'text'
      ? ctx.fontSizePx
      : ctx.fontSizePx / scaleFor(ctx.level, ctx.consts);
  return { ...ctx, level, fontSizePx: baseEm * scaleFor(level, ctx.consts) };
}

function shiftOp(op: DrawOp, dx: number, dy: number): DrawOp {
  if (op.type === 'stroke') {
    return { ...op, points: op.points.map((p) => ({ x: p.x + dx, y: p.y + dy })) };
  }
  return { ...op, x: op.x + dx, y: op.y + dy };
}

function shiftBox(b: MathBox, dx: number, dy: number): MathBox {
  return { ...b, ops: b.ops.map((o) => shiftOp(o, dx, dy)) };
}

/** A single glyph's box, using real ink metrics when a measurer is available. */
function glyphBox(ch: string, style: MathStyle, ctx: MathLayoutCtx): MathBox {
  let width: number;
  let ascent: number;
  let descent: number;
  if (ctx.measureGlyph) {
    const m = ctx.measureGlyph(ch, ctx.fontSizePx, style);
    width = m.width;
    ascent = Math.max(0, m.ascent);
    descent = Math.max(0, m.descent);
  } else {
    const cp = ch.codePointAt(0) ?? 0;
    width = fu(ctx, ctx.font.advance(ctx.font.glyphForChar(cp)));
    ascent = fu(ctx, ctx.font.ascent);
    descent = fu(ctx, ctx.font.descent);
  }
  return {
    width,
    ascent,
    descent,
    ops: [{ type: 'glyph', text: ch, style, x: 0, y: 0, sizePx: ctx.fontSizePx }],
  };
}

function runAtoms(node: { text: string; style: MathStyle }, ctx: MathLayoutCtx): Atom[] {
  const atoms: Atom[] = [];
  for (const ch of node.text) {
    if (ch === ' ') {
      // explicit space: a thin-ish ord gap
      atoms.push({ box: { width: muToPx(4, ctx), ascent: 0, descent: 0, ops: [] }, cls: 'ord' });
      continue;
    }
    atoms.push({ box: glyphBox(ch, node.style, ctx), cls: classifyChar(ch) });
  }
  return atoms;
}

/** Reclassify Bin atoms that aren't really binary (TeX rule), then concatenate with spacing. */
function hlistAtoms(atoms: Atom[], ctx: MathLayoutCtx): MathBox {
  for (let i = 0; i < atoms.length; i++) {
    if (atoms[i].cls !== 'bin') continue;
    const prev = i > 0 ? atoms[i - 1].cls : null;
    const next = i < atoms.length - 1 ? atoms[i + 1].cls : null;
    if (prev === null || prev === 'bin' || prev === 'op' || prev === 'rel' || prev === 'open' || prev === 'punct') {
      atoms[i].cls = 'ord';
    } else if (next === 'rel' || next === 'close' || next === 'punct') {
      atoms[i].cls = 'ord';
    }
  }
  const suppress = ctx.level === 'script' || ctx.level === 'scriptScript';
  let width = 0;
  let ascent = 0;
  let descent = 0;
  const ops: DrawOp[] = [];
  for (let i = 0; i < atoms.length; i++) {
    if (i > 0) {
      let mu = SPACE[atoms[i - 1].cls]?.[atoms[i].cls] ?? 0;
      if (suppress && mu > 3) mu = 0; // script styles drop medium/thick spaces
      width += muToPx(mu, ctx);
    }
    const b = atoms[i].box;
    for (const op of b.ops) ops.push(shiftOp(op, width, 0));
    width += b.width;
    ascent = Math.max(ascent, b.ascent);
    descent = Math.max(descent, b.descent);
  }
  return { width, ascent, descent, ops };
}

/** Concatenate boxes left-to-right with NO inter-atom spacing (internal combine). */
function hjoin(boxes: MathBox[]): MathBox {
  let width = 0;
  let ascent = 0;
  let descent = 0;
  const ops: DrawOp[] = [];
  for (const b of boxes) {
    for (const op of b.ops) ops.push(shiftOp(op, width, 0));
    width += b.width;
    ascent = Math.max(ascent, b.ascent);
    descent = Math.max(descent, b.descent);
  }
  return { width, ascent, descent, ops };
}

export function layoutMath(nodes: MathNode[], ctx: MathLayoutCtx): MathBox {
  return hlistAtoms(nodes.flatMap((n) => layoutAtoms(n, ctx)), ctx);
}

function layoutAtoms(node: MathNode, ctx: MathLayoutCtx): Atom[] {
  switch (node.kind) {
    case 'run':
      return runAtoms(node, ctx);
    case 'group':
      return [{ box: layoutMath(node.items, ctx), cls: 'ord' }];
    case 'fraction':
      return [{ box: fractionBox(node, ctx), cls: 'inner' }];
    case 'sup':
    case 'sub':
    case 'subSup':
      return [{ box: scriptBox(node, ctx), cls: scriptClass(node) }];
    case 'nary':
      return [{ box: naryOpBox(node, ctx), cls: 'op' }, ...node.body.flatMap((n) => layoutAtoms(n, ctx))];
    case 'delimiter':
      return [{ box: delimiterBox(node, ctx), cls: 'inner' }];
    case 'radical':
      return [{ box: radicalBox(node, ctx), cls: 'ord' }];
    case 'func':
      return [{ box: funcBox(node, ctx), cls: 'op' }];
  }
}

/** A script keeps the atom class of its base's leading character (so e.g. ∑ⁿ is Op). */
function scriptClass(node: Extract<MathNode, { kind: 'sup' | 'sub' | 'subSup' }>): AtomClass {
  const first = node.base[0];
  if (first?.kind === 'run' && first.text) return classifyChar([...first.text][0]);
  return 'ord';
}

function fractionBox(node: Extract<MathNode, { kind: 'fraction' }>, ctx: MathLayoutCtx): MathBox {
  const c = ctx.consts;
  const num = layoutMath(node.num, ctx);
  const den = layoutMath(node.den, ctx);
  const axis = fu(ctx, c.axisHeight);
  const rule = node.bar === false ? 0 : fu(ctx, c.fractionRuleThickness);
  const gapNum = fu(ctx, c.fractionNumeratorGapMin);
  const gapDen = fu(ctx, c.fractionDenominatorGapMin);
  const width = Math.max(num.width, den.width);

  let shiftUp = fu(ctx, c.fractionNumeratorShiftUp);
  shiftUp = Math.max(shiftUp, axis + rule / 2 + gapNum + num.descent);
  let shiftDown = fu(ctx, c.fractionDenominatorShiftDown);
  shiftDown = Math.max(shiftDown, den.ascent + rule / 2 + gapDen - axis);

  const ops: DrawOp[] = [
    ...shiftBox(num, (width - num.width) / 2, -shiftUp).ops,
    ...shiftBox(den, (width - den.width) / 2, shiftDown).ops,
  ];
  if (rule > 0) ops.push({ type: 'rule', x: 0, y: -(axis + rule / 2), w: width, h: rule });

  return { width, ascent: shiftUp + num.ascent, descent: shiftDown + den.descent, ops };
}

function scriptBox(
  node: Extract<MathNode, { kind: 'sup' | 'sub' | 'subSup' }>,
  ctx: MathLayoutCtx,
): MathBox {
  const c = ctx.consts;
  const base = layoutMath(node.base, ctx);
  const sctx = scriptCtx(ctx);
  const gap = fu(ctx, c.spaceAfterScript);
  const ops: DrawOp[] = [...base.ops];
  let ascent = base.ascent;
  let descent = base.descent;
  let scriptWidth = 0;

  if (node.sup) {
    const sup = layoutMath(node.sup, sctx);
    const shiftUp = Math.max(
      fu(ctx, c.superscriptShiftUp),
      base.ascent - fu(ctx, c.superscriptBottomMin), // ride near the top of tall bases
      sup.descent + fu(ctx, c.superscriptBottomMin),
    );
    ops.push(...shiftBox(sup, base.width, -shiftUp).ops);
    ascent = Math.max(ascent, shiftUp + sup.ascent);
    scriptWidth = Math.max(scriptWidth, sup.width);
  }
  if (node.sub) {
    const sub = layoutMath(node.sub, sctx);
    const shiftDown = Math.max(
      fu(ctx, c.subscriptShiftDown),
      sub.ascent - fu(ctx, c.subscriptTopMax),
    );
    ops.push(...shiftBox(sub, base.width, shiftDown).ops);
    descent = Math.max(descent, shiftDown + sub.descent);
    scriptWidth = Math.max(scriptWidth, sub.width);
  }

  return { width: base.width + scriptWidth + gap, ascent, descent, ops };
}

function naryOpBox(node: Extract<MathNode, { kind: 'nary' }>, ctx: MathLayoutCtx): MathBox {
  const opNode: MathNode = { kind: 'run', text: node.op, style: 'roman' };
  const scripted: MathNode =
    node.sup && node.sub
      ? { kind: 'subSup', base: [opNode], sup: node.sup, sub: node.sub }
      : node.sup
        ? { kind: 'sup', base: [opNode], sup: node.sup }
        : node.sub
          ? { kind: 'sub', base: [opNode], sub: node.sub }
          : opNode;
  // Build just the operator + its limits (the body is laid out as sibling atoms).
  return hlistAtoms(layoutAtoms(scripted, ctx), ctx);
}

function funcBox(node: Extract<MathNode, { kind: 'func' }>, ctx: MathLayoutCtx): MathBox {
  const name = layoutMath(node.name, ctx);
  const thin = { width: muToPx(3, ctx), ascent: 0, descent: 0, ops: [] as DrawOp[] };
  const arg = layoutMath(node.arg, ctx);
  return hjoin([name, thin, arg]);
}

function radicalBox(node: Extract<MathNode, { kind: 'radical' }>, ctx: MathLayoutCtx): MathBox {
  const c = ctx.consts;
  const radicand = layoutMath(node.radicand, ctx);
  const rule = fu(ctx, c.radicalRuleThickness);
  const gap = fu(ctx, c.radicalVerticalGap);
  const extra = fu(ctx, c.radicalExtraAscender);

  // Phase 1: synthesize the surd as a stroked polyline sized to the radicand. This
  // avoids depending on the font's U+221A glyph (which descends below the baseline
  // and needs per-glyph extents we don't parse until Phase 2) and always fits.
  const top = -(radicand.ascent + gap + rule);
  const bottom = radicand.descent;
  const h = bottom - top;
  const surdW = Math.max(fu(ctx, ctx.font.unitsPerEm) * 0.055, h * 0.45);
  const kern = rule * 1.5;
  const radX = surdW + kern;
  const vinculumEnd = radX + radicand.width;

  const surd: DrawOp = {
    type: 'stroke',
    lineWidth: rule,
    points: [
      { x: 0, y: top + h * 0.55 },
      { x: surdW * 0.3, y: top + h * 0.42 },
      { x: surdW * 0.52, y: bottom },
      { x: surdW * 0.85, y: top + rule / 2 },
      { x: vinculumEnd, y: top + rule / 2 },
    ],
  };

  const ops: DrawOp[] = [surd, ...shiftBox(radicand, radX, 0).ops];
  return {
    width: vinculumEnd,
    ascent: radicand.ascent + gap + rule + extra,
    descent: radicand.descent,
    ops,
  };
}

function delimiterBox(node: Extract<MathNode, { kind: 'delimiter' }>, ctx: MathLayoutCtx): MathBox {
  // Phase 1: fixed-size delimiters (stretchy variants are Phase 2).
  const beg = glyphBox(node.begChar || '(', 'roman', ctx);
  const inner = node.items.map((g, i) =>
    i === 0 ? layoutMath(g, ctx) : hjoin([glyphBox(',', 'roman', ctx), layoutMath(g, ctx)]),
  );
  const end = glyphBox(node.endChar || ')', 'roman', ctx);
  return hjoin([beg, ...inner, end]);
}
