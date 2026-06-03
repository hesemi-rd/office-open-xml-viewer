import type { MathConstants } from './math-table';
import type { MathFont } from './font';
import type { MathNode, MathStyle } from '../types/math';

export type MathLevel = 'display' | 'text' | 'script' | 'scriptScript';

export interface MathLayoutCtx {
  font: MathFont;
  consts: MathConstants;
  /** em size in CSS px at this level. */
  fontSizePx: number;
  level: MathLevel;
}

export type DrawOp =
  | { type: 'glyph'; text: string; style: MathStyle; x: number; y: number; sizePx: number }
  | { type: 'rule'; x: number; y: number; w: number; h: number };

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
  // Child em size scales relative to the *root* text size, not cumulatively per call,
  // so derive it from the level's scale factor against the base em.
  const baseEm = ctx.level === 'display' || ctx.level === 'text'
    ? ctx.fontSizePx
    : ctx.fontSizePx / scaleFor(ctx.level, ctx.consts);
  return { ...ctx, level, fontSizePx: baseEm * scaleFor(level, ctx.consts) };
}

function shiftOp(op: DrawOp, dx: number, dy: number): DrawOp {
  return { ...op, x: op.x + dx, y: op.y + dy };
}

function shiftBox(b: MathBox, dx: number, dy: number): MathBox {
  return { ...b, ops: b.ops.map((o) => shiftOp(o, dx, dy)) };
}

function runBox(node: { text: string; style: MathStyle }, ctx: MathLayoutCtx): MathBox {
  const { font } = ctx;
  let width = 0;
  for (const ch of node.text) {
    width += fu(ctx, font.advance(font.glyphForChar(ch.codePointAt(0)!)));
  }
  return {
    width,
    ascent: fu(ctx, font.ascent),
    descent: fu(ctx, font.descent),
    ops: [{ type: 'glyph', text: node.text, style: node.style, x: 0, y: 0, sizePx: ctx.fontSizePx }],
  };
}

/** Concatenate boxes left-to-right sharing a baseline (Phase 1: uniform spacing). */
function hlist(boxes: MathBox[]): MathBox {
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
  return hlist(nodes.map((n) => layoutNode(n, ctx)));
}

function layoutNode(node: MathNode, ctx: MathLayoutCtx): MathBox {
  switch (node.kind) {
    case 'run':
      return runBox(node, ctx);
    case 'group':
      return layoutMath(node.items, ctx);
    case 'fraction':
      return fractionBox(node, ctx);
    case 'sup':
    case 'sub':
    case 'subSup':
      return scriptBox(node, ctx);
    case 'nary':
      return naryBox(node, ctx);
    case 'delimiter':
      return delimiterBox(node, ctx);
    case 'radical':
      return radicalBox(node, ctx);
    case 'func':
      return hlist([
        layoutMath(node.name, ctx),
        runBox({ text: ' ', style: 'roman' }, ctx),
        layoutMath(node.arg, ctx),
      ]);
  }
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

  // Work in upward-positive coords (baseline = 0, up = +), then convert to canvas y.
  let shiftUp = fu(ctx, c.fractionNumeratorShiftUp);
  shiftUp = Math.max(shiftUp, axis + rule / 2 + gapNum + num.descent);
  let shiftDown = fu(ctx, c.fractionDenominatorShiftDown);
  shiftDown = Math.max(shiftDown, den.ascent + rule / 2 + gapDen - axis);

  const ops: DrawOp[] = [
    ...shiftBox(num, (width - num.width) / 2, -shiftUp).ops,
    ...shiftBox(den, (width - den.width) / 2, shiftDown).ops,
  ];
  if (rule > 0) ops.push({ type: 'rule', x: 0, y: -(axis + rule / 2), w: width, h: rule });

  return {
    width,
    ascent: shiftUp + num.ascent,
    descent: shiftDown + den.descent,
    ops,
  };
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

function naryBox(node: Extract<MathNode, { kind: 'nary' }>, ctx: MathLayoutCtx): MathBox {
  // Phase 1: operator at text size, limits attached as scripts (non-stretchy).
  const opNode: MathNode = { kind: 'run', text: node.op, style: 'roman' };
  const scripted: MathNode =
    node.sup && node.sub
      ? { kind: 'subSup', base: [opNode], sup: node.sup, sub: node.sub }
      : node.sup
        ? { kind: 'sup', base: [opNode], sup: node.sup }
        : node.sub
          ? { kind: 'sub', base: [opNode], sub: node.sub }
          : opNode;
  return hlist([
    layoutNode(scripted, ctx),
    runBox({ text: ' ', style: 'roman' }, ctx),
    layoutMath(node.body, ctx),
  ]);
}

function radicalBox(node: Extract<MathNode, { kind: 'radical' }>, ctx: MathLayoutCtx): MathBox {
  const c = ctx.consts;
  const radicand = layoutMath(node.radicand, ctx);
  // Phase 1: fixed-size surd glyph (U+221A). Stretchy variants are Phase 2.
  const surd = runBox({ text: '√', style: 'roman' }, ctx);
  const rule = fu(ctx, c.radicalRuleThickness);
  const gap = fu(ctx, c.radicalVerticalGap);
  const extra = fu(ctx, c.radicalExtraAscender);
  const kern = fu(ctx, c.fractionRuleThickness); // small horizontal breathing room after surd

  const radX = surd.width + kern;
  // Vinculum (overline) sits gap above the radicand top; its top adds the rule thickness.
  const vinculumTopY = -(radicand.ascent + gap + rule);
  const ops: DrawOp[] = [
    ...surd.ops, // surd at baseline origin
    ...shiftBox(radicand, radX, 0).ops,
    { type: 'rule', x: radX, y: vinculumTopY, w: radicand.width, h: rule },
    // extend the vinculum left to meet the top of the surd
    { type: 'rule', x: surd.width - kern * 0.5, y: vinculumTopY, w: radX - (surd.width - kern * 0.5), h: rule },
  ];

  return {
    width: radX + radicand.width,
    ascent: Math.max(surd.ascent, radicand.ascent + gap + rule + extra),
    descent: Math.max(surd.descent, radicand.descent),
    ops,
  };
}

function delimiterBox(node: Extract<MathNode, { kind: 'delimiter' }>, ctx: MathLayoutCtx): MathBox {
  // Phase 1: fixed-size delimiters via fillText (stretchy variants are Phase 2).
  const beg = runBox({ text: node.begChar || '(', style: 'roman' }, ctx);
  const inner = node.items.map((g, i) =>
    i === 0
      ? layoutMath(g, ctx)
      : hlist([runBox({ text: ',', style: 'roman' }, ctx), layoutMath(g, ctx)]),
  );
  const end = runBox({ text: node.endChar || ')', style: 'roman' }, ctx);
  return hlist([beg, ...inner, end]);
}
