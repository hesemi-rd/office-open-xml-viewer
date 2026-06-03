# OOXML Math Phase 1 — Core Engine + docx Inline Rendering

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render OMML equations from docx onto Canvas as crisp vectors via a new shared `packages/core/src/math/` engine, no DOM, deterministic in node.

**Architecture:** Each Rust parser extracts `<m:oMath>` into a shared OMML AST (TS-typed JSON). `core/src/math/` parses an OpenType math font (Latin Modern Math) for metrics + MATH-table constants, lays the AST out into a positioned box tree using TeX-style rules, and draws the boxes to a Canvas2D context. Layout uses font-parsed metrics so it is fully deterministic and unit-testable without a canvas; drawing uses `fillText` (Phase 1) with the registered font.

**Tech Stack:** TypeScript (core engine, zero runtime deps), Rust/serde (parser extraction), vitest (unit tests), Latin Modern Math (OFL math font).

**Scope note:** Phase 1 covers fractions, sub/superscripts, n-ary operators (non-stretchy), ordinary symbols/Greek, operators, and correct inter-atom spacing — enough for the bulk of `private/sample-6.docx`. Stretchy/large radicals & delimiters (CFF outline + MathVariants/GlyphAssembly), accents, matrices, and pptx/xlsx extraction are Phase 2/3 (separate plans).

---

## File Structure

- `packages/core/src/types/math.ts` — OMML AST types (the contract). Pure types + a tiny node-builder for tests.
- `packages/core/src/math/math-table.ts` — minimal OpenType MATH table reader (MathConstants, italic corrections).
- `packages/core/src/math/font.ts` — parse a math font ArrayBuffer: `unitsPerEm`, cmap (char→gid), hmtx (gid→advance), glyph vertical extents; lazy load + cache; integrate MATH table.
- `packages/core/src/math/layout.ts` — `layoutMath(ast, font, style): MathBox` (AST → positioned box tree).
- `packages/core/src/math/render.ts` — `renderMathBox(ctx, box, x, baseline)` and `measureMath`.
- `packages/core/src/math/index.ts` — public surface re-exports.
- `packages/core/src/index.ts` — add math exports.
- `packages/core/assets/LatinModernMath.otf` — bundled math font (OFL); plus `LICENSE` note.
- `packages/docx/parser/src/types.rs` — add `Math` variant to `DocRun` + OMML structs.
- `packages/docx/parser/src/parser.rs` — extract `m:oMath` / `m:oMathPara`.
- `packages/docx/src/types.ts` — mirror the math run type.
- `packages/docx/src/renderer.ts` — measure + draw math segments.

---

## Task 1: OMML AST types

**Files:**
- Create: `packages/core/src/types/math.ts`
- Test: `packages/core/src/types/math.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/types/math.test.ts
import { describe, it, expect } from 'vitest';
import { isMathNode, type MathNode } from './math';

describe('OMML AST', () => {
  it('discriminates node kinds', () => {
    const frac: MathNode = {
      kind: 'fraction',
      num: [{ kind: 'run', text: '1', style: 'italic' }],
      den: [{ kind: 'run', text: 'x', style: 'italic' }],
    };
    expect(frac.kind).toBe('fraction');
    expect(isMathNode(frac)).toBe(true);
    expect(isMathNode({ foo: 1 })).toBe(false);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- packages/core/src/types/math.test.ts`
Expected: FAIL — `Cannot find module './math'`.

- [ ] **Step 3: Write the types**

```ts
// packages/core/src/types/math.ts
// OMML AST shared by docx/xlsx/pptx parsers. Mirrors ECMA-376 §22.1.2.
export type MathStyle = 'roman' | 'italic' | 'bold' | 'boldItalic';

export interface MathRun { kind: 'run'; text: string; style: MathStyle }
export interface MathFraction { kind: 'fraction'; num: MathNode[]; den: MathNode[]; bar?: boolean }
export interface MathScript {
  kind: 'sup' | 'sub' | 'subSup';
  base: MathNode[];
  sup?: MathNode[];
  sub?: MathNode[];
}
export interface MathNary {
  kind: 'nary';
  op: string;            // e.g. '∑', '∫', '∏'
  sub?: MathNode[];
  sup?: MathNode[];
  body: MathNode[];
}
export interface MathDelimiter {
  kind: 'delimiter';
  begChar: string;       // default '('
  endChar: string;       // default ')'
  items: MathNode[][];   // separated groups
}
export interface MathFunc { kind: 'func'; name: MathNode[]; arg: MathNode[] }
export interface MathGroup { kind: 'group'; items: MathNode[] }

export type MathNode =
  | MathRun | MathFraction | MathScript | MathNary | MathDelimiter | MathFunc | MathGroup;

const KINDS = new Set(['run', 'fraction', 'sup', 'sub', 'subSup', 'nary', 'delimiter', 'func', 'group']);
export function isMathNode(v: unknown): v is MathNode {
  return !!v && typeof v === 'object' && KINDS.has((v as { kind?: string }).kind ?? '');
}

/** Top-level math container as emitted by parsers. `display` = block (`m:oMathPara`). */
export interface MathFormula { nodes: MathNode[]; display: boolean }
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pnpm test -- packages/core/src/types/math.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/types/math.ts packages/core/src/types/math.test.ts
git commit -m "feat(core): add OMML math AST types"
```

---

## Task 2: Bundle the math font

**Files:**
- Create: `packages/core/assets/LatinModernMath.otf`
- Create: `packages/core/assets/LICENSE-LatinModernMath.txt`

- [ ] **Step 1: Download the OFL font from GUST (Latin Modern Math)**

```bash
cd packages/core && mkdir -p assets
curl -L -o /tmp/lm-math.zip https://www.gust.org.pl/projects/e-foundry/lm-math/download/latinmodern-math-1959.zip
unzip -o /tmp/lm-math.zip -d /tmp/lm-math
find /tmp/lm-math -name 'latinmodern-math.otf' -exec cp {} assets/LatinModernMath.otf \;
find /tmp/lm-math -iname 'GUST-FONT-LICENSE.txt' -exec cp {} assets/LICENSE-LatinModernMath.txt \;
ls -la assets/
```

Expected: `LatinModernMath.otf` (~390KB) and the license file present.

- [ ] **Step 2: Verify the font has a MATH table**

```bash
node -e "const b=require('fs').readFileSync('packages/core/assets/LatinModernMath.otf'); const n=b.readUInt16BE(4); let off=12,found=false; for(let i=0;i<n;i++){const tag=b.toString('ascii',off,off+4); if(tag==='MATH')found=true; off+=16;} console.log('MATH table present:', found);"
```

Expected: `MATH table present: true`.

- [ ] **Step 3: Commit**

```bash
git add packages/core/assets/LatinModernMath.otf packages/core/assets/LICENSE-LatinModernMath.txt
git commit -m "chore(core): bundle Latin Modern Math (OFL) for math rendering"
```

---

## Task 3: Math font parser (sfnt tables: head/cmap/hmtx)

**Files:**
- Create: `packages/core/src/math/font.ts`
- Test: `packages/core/src/math/font.test.ts`

Parse only what layout needs: `unitsPerEm` (head), char→glyphId (cmap format 4 + 12), glyphId→advance (hmtx + hhea numberOfHMetrics), ascent/descent (hhea or OS/2). All values returned in font units; callers scale by `fontSizePx / unitsPerEm`.

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/math/font.test.ts
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont, type MathFont } from './font';

let font: MathFont;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  font = parseMathFont(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
});

describe('parseMathFont', () => {
  it('reads unitsPerEm', () => {
    expect(font.unitsPerEm).toBe(1000); // LM Math is 1000 upm
  });
  it('maps ASCII and Greek to advances', () => {
    const x = font.glyphForChar('x'.codePointAt(0)!);
    expect(x).toBeGreaterThan(0);
    expect(font.advance(x)).toBeGreaterThan(0);
    const sum = font.glyphForChar('∑'.codePointAt(0)!);
    expect(sum).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- packages/core/src/math/font.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement the sfnt parser**

```ts
// packages/core/src/math/font.ts
export interface MathFont {
  unitsPerEm: number;
  ascent: number;   // font units
  descent: number;  // font units (positive magnitude)
  glyphForChar(cp: number): number;     // 0 = .notdef
  advance(gid: number): number;         // font units
  buffer: ArrayBuffer;                  // retained for MATH table + future outline use
  tableOffset(tag: string): number;     // sfnt table offset, -1 if absent
}

function tableDirectory(dv: DataView): Map<string, number> {
  const num = dv.getUint16(4);
  const map = new Map<string, number>();
  let off = 12;
  for (let i = 0; i < num; i++) {
    const tag = String.fromCharCode(dv.getUint8(off), dv.getUint8(off + 1), dv.getUint8(off + 2), dv.getUint8(off + 3));
    map.set(tag, dv.getUint32(off + 8));
    off += 16;
  }
  return map;
}

function parseCmap(dv: DataView, base: number): Map<number, number> {
  const numTables = dv.getUint16(base + 2);
  let best = -1;
  for (let i = 0; i < numTables; i++) {
    const rec = base + 4 + i * 8;
    const plat = dv.getUint16(rec), enc = dv.getUint16(rec + 2), sub = base + dv.getUint32(rec + 4);
    // Prefer Unicode BMP+full: (3,10) or (3,1) or (0,*)
    if ((plat === 3 && enc === 10) || (plat === 0 && enc >= 4)) return readSubtable(dv, sub);
    if ((plat === 3 && enc === 1) || plat === 0) best = sub;
  }
  return best >= 0 ? readSubtable(dv, best) : new Map();
}

function readSubtable(dv: DataView, off: number): Map<number, number> {
  const fmt = dv.getUint16(off);
  const m = new Map<number, number>();
  if (fmt === 4) {
    const segX2 = dv.getUint16(off + 6), segCount = segX2 / 2;
    const endO = off + 14, startO = endO + segX2 + 2, deltaO = startO + segX2, rangeO = deltaO + segX2;
    for (let s = 0; s < segCount; s++) {
      const end = dv.getUint16(endO + s * 2), start = dv.getUint16(startO + s * 2);
      const delta = dv.getUint16(deltaO + s * 2), ro = dv.getUint16(rangeO + s * 2);
      for (let c = start; c <= end && c !== 0xffff; c++) {
        let g: number;
        if (ro === 0) g = (c + delta) & 0xffff;
        else {
          const gi = dv.getUint16(rangeO + s * 2 + ro + (c - start) * 2);
          g = gi === 0 ? 0 : (gi + delta) & 0xffff;
        }
        if (g) m.set(c, g);
      }
    }
  } else if (fmt === 12) {
    const nGroups = dv.getUint32(off + 12);
    let g = off + 16;
    for (let i = 0; i < nGroups; i++) {
      const sc = dv.getUint32(g), ec = dv.getUint32(g + 4), sg = dv.getUint32(g + 8);
      for (let c = sc; c <= ec; c++) m.set(c, sg + (c - sc));
      g += 12;
    }
  }
  return m;
}

export function parseMathFont(buffer: ArrayBuffer): MathFont {
  const dv = new DataView(buffer);
  const dir = tableDirectory(dv);
  const head = dir.get('head')!, hhea = dir.get('hhea')!, hmtx = dir.get('hmtx')!, maxp = dir.get('maxp')!;
  const unitsPerEm = dv.getUint16(head + 18);
  const ascent = dv.getInt16(hhea + 4), descent = -dv.getInt16(hhea + 6);
  const numHMetrics = dv.getUint16(hhea + 34);
  const numGlyphs = dv.getUint16(maxp + 4);
  const cmap = parseCmap(dv, dir.get('cmap')!);

  const advance = (gid: number): number => {
    const i = gid < numHMetrics ? gid : numHMetrics - 1;
    return dv.getUint16(hmtx + i * 4);
  };
  return {
    unitsPerEm, ascent, descent, buffer,
    glyphForChar: (cp) => cmap.get(cp) ?? 0,
    advance: (gid) => (gid >= 0 && gid < numGlyphs ? advance(gid) : 0),
    tableOffset: (tag) => dir.get(tag) ?? -1,
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pnpm test -- packages/core/src/math/font.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/math/font.ts packages/core/src/math/font.test.ts
git commit -m "feat(core): parse math-font sfnt tables (head/cmap/hmtx)"
```

---

## Task 4: MATH table constants

**Files:**
- Create: `packages/core/src/math/math-table.ts`
- Test: `packages/core/src/math/math-table.test.ts`

Read `MathConstants` (the subset layout needs) from the MATH table. Each constant after the first few is a `MathValueRecord` (Int16 value + Uint16 deviceOffset); we take the value. See OpenType MATH spec.

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/math/math-table.test.ts
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont } from './font';
import { parseMathConstants, type MathConstants } from './math-table';

let mc: MathConstants;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  const font = parseMathFont(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  mc = parseMathConstants(font);
});

describe('parseMathConstants', () => {
  it('reads a positive axis height and fraction rule thickness', () => {
    expect(mc.axisHeight).toBeGreaterThan(0);
    expect(mc.fractionRuleThickness).toBeGreaterThan(0);
    expect(mc.scriptPercentScaleDown).toBeGreaterThan(0);
    expect(mc.scriptPercentScaleDown).toBeLessThan(100);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- packages/core/src/math/math-table.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement the MATH constants reader**

```ts
// packages/core/src/math/math-table.ts
import type { MathFont } from './font';

// All distances in font units (scale by fontSizePx/unitsPerEm). Percentages are 0..100.
export interface MathConstants {
  scriptPercentScaleDown: number;
  scriptScriptPercentScaleDown: number;
  axisHeight: number;
  fractionRuleThickness: number;
  fractionNumeratorShiftUp: number;
  fractionDenominatorShiftDown: number;
  fractionNumeratorGapMin: number;
  fractionDenominatorGapMin: number;
  superscriptShiftUp: number;
  subscriptShiftDown: number;
  superscriptBottomMin: number;
  subscriptTopMax: number;
  spaceAfterScript: number;
  upperLimitGapMin: number;
  lowerLimitGapMin: number;
  mathLeading: number;
}

export function parseMathConstants(font: MathFont): MathConstants {
  const off = font.tableOffset('MATH');
  if (off < 0) throw new Error('font has no MATH table');
  const dv = new DataView(font.buffer);
  const constsOff = off + dv.getUint16(off + 4); // MathConstants offset (after version 4 + 3 offsets... see layout below)
  // MATH header: version(4? no) — actual: MajorVersion(2)+MinorVersion(2)=4 bytes,
  // then MathConstantsOffset(2), MathGlyphInfoOffset(2), MathVariantsOffset(2).
  const c = off + dv.getUint16(off + 4);
  void constsOff;
  const u16 = (p: number) => dv.getUint16(c + p);
  const i16 = (p: number) => dv.getInt16(c + p);
  // First 4 fields are plain Int16 (percentages/values). After ScriptScriptPercentScaleDown
  // and DelimitedSubFormulaMinHeight(Uint16) etc., the rest are MathValueRecord (value=Int16 at start of record).
  // Field offsets per OpenType MATH MathConstants table layout:
  return {
    scriptPercentScaleDown: i16(0),
    scriptScriptPercentScaleDown: i16(2),
    // 4: DelimitedSubFormulaMinHeight (u16), 6: DisplayOperatorMinHeight (u16)
    mathLeading: i16(8),               // MathValueRecord.value
    axisHeight: i16(10),
    // 12 AccentBaseHeight, 14 FlattenedAccentBaseHeight,
    subscriptShiftDown: i16(16),
    subscriptTopMax: i16(18),
    // 20 SubscriptBaselineDropMin
    superscriptShiftUp: i16(22),
    // 24 SuperscriptShiftUpCramped
    superscriptBottomMin: i16(26),
    // 28 SuperscriptBaselineDropMax, 30 SubSuperscriptGapMin, 32 SuperscriptBottomMaxWithSubscript
    spaceAfterScript: i16(34),
    upperLimitGapMin: i16(36),
    lowerLimitGapMin: i16(38),
    // ... (40 Limits... ) up to fraction fields
    fractionNumeratorShiftUp: i16(50),
    // 52 FractionNumeratorDisplayStyleShiftUp
    fractionDenominatorShiftDown: i16(54),
    // 56 ...DisplayStyle
    fractionNumeratorGapMin: i16(58),
    // 60 ...DisplayStyleGapMin
    fractionRuleThickness: i16(62),
    fractionDenominatorGapMin: i16(64),
    u16: undefined as never,
  } as unknown as MathConstants;
}
```

> NOTE FOR IMPLEMENTER: the exact byte offsets above are the high-risk part. During TDD,
> assert each constant against a value dumped with a reference tool. Cross-check field
> ordering with the OpenType MATH `MathConstants` table definition
> (https://learn.microsoft.com/typography/opentype/spec/math#mathconstants-table). The
> first two fields are `Int16` percentages; `minConnectorOverlap` and several `*MinHeight`
> are `UInt16`; everything else is a `MathValueRecord` whose first 2 bytes are the `Int16`
> value. Fix offsets until the test asserting plausible ranges passes, then tighten the
> test to exact dumped values.

- [ ] **Step 4: Run test to verify it passes (iterating offsets as needed)**

Run: `pnpm test -- packages/core/src/math/math-table.test.ts`
Expected: PASS with plausible positive values.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/math/math-table.ts packages/core/src/math/math-table.test.ts
git commit -m "feat(core): read MATH table layout constants"
```

---

## Task 5: Layout engine — runs, fractions, scripts

**Files:**
- Create: `packages/core/src/math/layout.ts`
- Test: `packages/core/src/math/layout.test.ts`

Produce a positioned box tree. A `MathBox` has CSS-px metrics relative to its own origin: `width`, `ascent` (above baseline), `descent` (below baseline), and `children` with offsets. Style level (`'display'|'text'|'script'|'scriptScript'`) selects scale via `scriptPercentScaleDown`.

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/math/layout.test.ts
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont } from './font';
import { parseMathConstants } from './math-table';
import { layoutMath, type MathLayoutCtx } from './layout';
import type { MathNode } from '../types/math';

let ctx: MathLayoutCtx;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  const font = parseMathFont(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  ctx = { font, consts: parseMathConstants(font), fontSizePx: 12, level: 'text' };
});

describe('layoutMath', () => {
  it('lays out a single run with positive width and height', () => {
    const nodes: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const box = layoutMath(nodes, ctx);
    expect(box.width).toBeGreaterThan(0);
    expect(box.ascent).toBeGreaterThan(0);
  });
  it('stacks a fraction: total height exceeds numerator alone', () => {
    const num: MathNode[] = [{ kind: 'run', text: '1', style: 'italic' }];
    const den: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const single = layoutMath(num, ctx);
    const frac = layoutMath([{ kind: 'fraction', num, den }], ctx);
    expect(frac.ascent + frac.descent).toBeGreaterThan(single.ascent + single.descent);
    // numerator sits above the math axis, denominator below
    expect(frac.ascent).toBeGreaterThan(0);
    expect(frac.descent).toBeGreaterThan(0);
  });
  it('raises a superscript above the base ascent', () => {
    const base: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const plain = layoutMath(base, ctx);
    const sup = layoutMath([{ kind: 'sup', base, sup: [{ kind: 'run', text: '2', style: 'italic' }] }], ctx);
    expect(sup.ascent).toBeGreaterThan(plain.ascent);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- packages/core/src/math/layout.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement layout (runs, horizontal lists, fraction, scripts, nary, delimiter, func, group)**

```ts
// packages/core/src/math/layout.ts
import type { MathConstants } from './math-table';
import type { MathFont } from './font';
import type { MathNode, MathStyle } from '../types/math';

export type MathLevel = 'display' | 'text' | 'script' | 'scriptScript';
export interface MathLayoutCtx {
  font: MathFont;
  consts: MathConstants;
  fontSizePx: number;   // em size at this level
  level: MathLevel;
}

export type DrawOp =
  | { type: 'glyph'; text: string; style: MathStyle; x: number; y: number; sizePx: number }
  | { type: 'rule'; x: number; y: number; w: number; h: number };

export interface MathBox {
  width: number;
  ascent: number;   // above baseline (px, positive up)
  descent: number;  // below baseline (px, positive down)
  ops: DrawOp[];    // positioned relative to this box's baseline origin (x=0, baseline y=0)
}

const fu = (ctx: MathLayoutCtx, v: number) => (v * ctx.fontSizePx) / ctx.font.unitsPerEm;

function scaleFor(level: MathLevel, c: MathConstants): number {
  if (level === 'script') return c.scriptPercentScaleDown / 100;
  if (level === 'scriptScript') return (c.scriptPercentScaleDown / 100) * (c.scriptScriptPercentScaleDown / 100);
  return 1;
}
function childLevel(level: MathLevel): MathLevel {
  return level === 'display' || level === 'text' ? 'script' : 'scriptScript';
}

function runBox(node: { text: string; style: MathStyle }, ctx: MathLayoutCtx): MathBox {
  const { font } = ctx;
  let width = 0;
  for (const ch of node.text) width += fu(ctx, font.advance(font.glyphForChar(ch.codePointAt(0)!)));
  return {
    width,
    ascent: fu(ctx, font.ascent),
    descent: fu(ctx, font.descent),
    ops: [{ type: 'glyph', text: node.text, style: node.style, x: 0, y: 0, sizePx: ctx.fontSizePx }],
  };
}

/** Concatenate boxes left-to-right sharing a baseline. (Phase 1: uniform spacing; refine per atom class later.) */
function hlist(boxes: MathBox[]): MathBox {
  let width = 0, ascent = 0, descent = 0;
  const ops: DrawOp[] = [];
  for (const b of boxes) {
    for (const op of b.ops) ops.push(shiftOp(op, width, 0));
    width += b.width;
    ascent = Math.max(ascent, b.ascent);
    descent = Math.max(descent, b.descent);
  }
  return { width, ascent, descent, ops };
}

function shiftOp(op: DrawOp, dx: number, dy: number): DrawOp {
  return op.type === 'glyph'
    ? { ...op, x: op.x + dx, y: op.y + dy }
    : { ...op, x: op.x + dx, y: op.y + dy };
}
function shiftBox(b: MathBox, dx: number, dy: number): MathBox {
  return { ...b, ops: b.ops.map((o) => shiftOp(o, dx, dy)) };
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
    case 'func':
      return hlist([layoutMath(node.name, ctx), runBox({ text: ' ', style: 'roman' }, ctx), layoutMath(node.arg, ctx)]);
  }
}

function fractionBox(node: Extract<MathNode, { kind: 'fraction' }>, ctx: MathLayoutCtx): MathBox {
  const c = ctx.consts;
  const num = layoutMath(node.num, ctx);
  const den = layoutMath(node.den, ctx);
  const axis = fu(ctx, c.axisHeight);
  const rule = node.bar === false ? 0 : fu(ctx, c.fractionRuleThickness);
  const width = Math.max(num.width, den.width);
  const numShift = fu(ctx, c.fractionNumeratorShiftUp);
  const denShift = fu(ctx, c.fractionDenominatorShiftDown);
  // numerator baseline placed so its descent clears the rule by the gap
  const numBaselineY = -(axis + rule / 2 + fu(ctx, c.fractionNumeratorGapMin) + num.descent);
  const denBaselineY = axis - rule / 2 + fu(ctx, c.fractionDenominatorGapMin) + den.ascent;
  void numShift; void denShift; // shifts refined during TDD against references
  const ops: DrawOp[] = [
    ...shiftBox(num, (width - num.width) / 2, numBaselineY).ops,
    ...shiftBox(den, (width - den.width) / 2, denBaselineY).ops,
  ];
  if (rule > 0) ops.push({ type: 'rule', x: 0, y: -(axis), w: width, h: rule });
  const ascent = -numBaselineY + num.ascent;
  const descent = denBaselineY + den.descent;
  return { width, ascent, descent, ops };
}

function scriptBox(node: Extract<MathNode, { kind: 'sup' | 'sub' | 'subSup' }>, ctx: MathLayoutCtx): MathBox {
  const c = ctx.consts;
  const base = layoutMath(node.base, ctx);
  const sctx: MathLayoutCtx = { ...ctx, level: childLevel(ctx.level), fontSizePx: ctx.fontSizePx * scaleFor(childLevel(ctx.level), c) };
  const ops: DrawOp[] = [...base.ops];
  let width = base.width, ascent = base.ascent, descent = base.descent;
  const gap = fu(ctx, c.spaceAfterScript);
  if (node.sup) {
    const sup = layoutMath(node.sup, sctx);
    const shiftUp = Math.max(fu(ctx, c.superscriptShiftUp), base.ascent - sup.descent + fu(ctx, c.superscriptBottomMin));
    ops.push(...shiftBox(sup, base.width, -shiftUp).ops);
    width = Math.max(width, base.width + sup.width + gap);
    ascent = Math.max(ascent, shiftUp + sup.ascent);
  }
  if (node.sub) {
    const sub = layoutMath(node.sub, sctx);
    const shiftDown = Math.max(fu(ctx, c.subscriptShiftDown), sub.ascent - base.descent + 0);
    ops.push(...shiftBox(sub, base.width, shiftDown).ops);
    width = Math.max(width, base.width + sub.width + gap);
    descent = Math.max(descent, shiftDown + sub.descent);
  }
  return { width, ascent, descent, ops };
}

function naryBox(node: Extract<MathNode, { kind: 'nary' }>, ctx: MathLayoutCtx): MathBox {
  // Phase 1: operator drawn at text size with sub/sup as scripts (non-stretchy).
  const opNode: MathNode = { kind: 'run', text: node.op, style: 'roman' };
  const scripted: MathNode =
    node.sup && node.sub ? { kind: 'subSup', base: [opNode], sup: node.sup, sub: node.sub }
    : node.sup ? { kind: 'sup', base: [opNode], sup: node.sup }
    : node.sub ? { kind: 'sub', base: [opNode], sub: node.sub }
    : opNode;
  return hlist([layoutNode(scripted, ctx), runBox({ text: ' ', style: 'roman' }, ctx), layoutMath(node.body, ctx)]);
}

function delimiterBox(node: Extract<MathNode, { kind: 'delimiter' }>, ctx: MathLayoutCtx): MathBox {
  // Phase 1: fixed-size delimiters via fillText (stretchy variants are Phase 2).
  const beg = runBox({ text: node.begChar || '(', style: 'roman' }, ctx);
  const inner = node.items.map((g, i) =>
    i === 0 ? layoutMath(g, ctx) : hlist([runBox({ text: '∣', style: 'roman' }, ctx), layoutMath(g, ctx)]));
  const end = runBox({ text: node.endChar || ')', style: 'roman' }, ctx);
  return hlist([beg, ...inner, end]);
}
```

> NOTE FOR IMPLEMENTER: the fraction/script shift formulas are first-cut TeX-style
> approximations. The tests assert *relationships* (fraction taller than numerator,
> superscript raises ascent), not absolute pixels, so they stay green while you refine
> constants against the VRT in Task 8. Do not hardcode magic pixel constants — derive
> everything from `MathConstants` (CLAUDE.md spec-first rule).

- [ ] **Step 4: Run test to verify it passes**

Run: `pnpm test -- packages/core/src/math/layout.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/math/layout.ts packages/core/src/math/layout.test.ts
git commit -m "feat(core): math layout engine (runs/fractions/scripts/nary/delimiters)"
```

---

## Task 6: Canvas renderer + public API

**Files:**
- Create: `packages/core/src/math/render.ts`
- Create: `packages/core/src/math/index.ts`
- Modify: `packages/core/src/index.ts`
- Test: `packages/core/src/math/render.test.ts`

- [ ] **Step 1: Write the failing test (mock ctx records calls)**

```ts
// packages/core/src/math/render.test.ts
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont } from './font';
import { parseMathConstants } from './math-table';
import { layoutMath, type MathLayoutCtx } from './layout';
import { renderMathBox } from './render';
import type { MathNode } from '../types/math';

let ctx: MathLayoutCtx;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  const font = parseMathFont(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  ctx = { font, consts: parseMathConstants(font), fontSizePx: 16, level: 'text' };
});

function mockCtx() {
  const calls: string[] = [];
  return {
    calls,
    save() { calls.push('save'); },
    restore() { calls.push('restore'); },
    fillText(t: string, x: number, y: number) { calls.push(`text:${t}@${x.toFixed(1)},${y.toFixed(1)}`); },
    fillRect(x: number, y: number, w: number, h: number) { calls.push(`rect:${w.toFixed(1)}x${h.toFixed(1)}`); },
    set font(_v: string) {}, get font() { return ''; },
    set fillStyle(_v: string) {}, get fillStyle() { return ''; },
  } as unknown as CanvasRenderingContext2D & { calls: string[] };
}

describe('renderMathBox', () => {
  it('draws glyphs and a fraction rule', () => {
    const nodes: MathNode[] = [{ kind: 'fraction', num: [{ kind: 'run', text: '1', style: 'italic' }], den: [{ kind: 'run', text: 'x', style: 'italic' }] }];
    const box = layoutMath(nodes, ctx);
    const m = mockCtx();
    renderMathBox(m, box, 10, 100, '#000', 'LatinModernMath');
    expect((m as unknown as { calls: string[] }).calls.some((c) => c.startsWith('text:1'))).toBe(true);
    expect((m as unknown as { calls: string[] }).calls.some((c) => c.startsWith('rect:'))).toBe(true);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- packages/core/src/math/render.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement renderer + index**

```ts
// packages/core/src/math/render.ts
import type { MathBox, DrawOp } from './layout';

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

function styleToFont(style: string, sizePx: number, family: string): string {
  const bold = style === 'bold' || style === 'boldItalic';
  const italic = style === 'italic' || style === 'boldItalic';
  return `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px "${family}"`;
}

/** Draw a laid-out math box. `(x, baseline)` is the box origin (baseline y in canvas px). */
export function renderMathBox(ctx: Ctx2D, box: MathBox, x: number, baseline: number, color: string, family: string): void {
  ctx.save();
  ctx.fillStyle = color;
  for (const op of box.ops) drawOp(ctx, op, x, baseline, family);
  ctx.restore();
}

function drawOp(ctx: Ctx2D, op: DrawOp, x: number, baseline: number, family: string): void {
  if (op.type === 'glyph') {
    ctx.font = styleToFont(op.style, op.sizePx, family);
    ctx.fillText(op.text, x + op.x, baseline + op.y);
  } else {
    ctx.fillRect(x + op.x, baseline + op.y, op.w, op.h);
  }
}

export interface MathMetrics { width: number; ascent: number; descent: number }
export function measureMathBox(box: MathBox): MathMetrics {
  return { width: box.width, ascent: box.ascent, descent: box.descent };
}
```

```ts
// packages/core/src/math/index.ts
export { layoutMath, type MathBox, type MathLayoutCtx, type MathLevel, type DrawOp } from './layout';
export { renderMathBox, measureMathBox, type MathMetrics } from './render';
export { parseMathFont, type MathFont } from './font';
export { parseMathConstants, type MathConstants } from './math-table';
```

```ts
// packages/core/src/index.ts  — append:
export * from './math';
export type { MathNode, MathFormula, MathStyle } from './types/math';
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pnpm test -- packages/core/src/math/render.test.ts`
Expected: PASS.

- [ ] **Step 5: Typecheck + commit**

```bash
pnpm --filter @silurus/ooxml-core typecheck
git add packages/core/src/math packages/core/src/index.ts
git commit -m "feat(core): canvas math renderer + public API"
```

---

## Task 7: docx parser — extract OMML

**Files:**
- Modify: `packages/docx/parser/src/types.rs` (add `Math` variant + OMML structs)
- Modify: `packages/docx/parser/src/parser.rs` (parse `m:oMath` / `m:oMathPara`)
- Test: add a Rust unit test in `packages/docx/parser/src/parser.rs`

- [ ] **Step 1: Add serde types mirroring the TS AST**

In `types.rs`, add (tag/style names must serialize to match `core/src/types/math.ts`):

```rust
#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum MathNode {
    #[serde(rename = "run")]
    Run { text: String, style: String },
    #[serde(rename = "fraction")]
    Fraction { num: Vec<MathNode>, den: Vec<MathNode> },
    #[serde(rename = "sup")]
    Sup { base: Vec<MathNode>, sup: Vec<MathNode> },
    #[serde(rename = "sub")]
    Sub { base: Vec<MathNode>, sub: Vec<MathNode> },
    #[serde(rename = "subSup")]
    SubSup { base: Vec<MathNode>, sup: Vec<MathNode>, sub: Vec<MathNode> },
    #[serde(rename = "nary")]
    Nary { op: String, #[serde(skip_serializing_if = "Vec::is_empty")] sub: Vec<MathNode>, #[serde(skip_serializing_if = "Vec::is_empty")] sup: Vec<MathNode>, body: Vec<MathNode> },
    #[serde(rename = "delimiter")]
    Delimiter { beg_char: String, end_char: String, items: Vec<Vec<MathNode>> },
    #[serde(rename = "func")]
    Func { name: Vec<MathNode>, arg: Vec<MathNode> },
    #[serde(rename = "group")]
    Group { items: Vec<MathNode> },
}
```

Add a `DocRun::Math` variant (matching the existing `#[serde(tag = ...)]` pattern on `DocRun`, see `types.rs:282`):

```rust
    Math { nodes: Vec<MathNode>, display: bool },
```

- [ ] **Step 2: Write the failing Rust test**

```rust
// in parser.rs (#[cfg(test)] mod)
#[test]
fn parses_simple_fraction() {
    let xml = r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num>
      <m:den><m:r><m:t>x</m:t></m:r></m:den></m:f></m:oMath></w:p>"#;
    let runs = parse_paragraph_runs_for_test(xml);
    assert!(runs.iter().any(|r| matches!(r, DocRun::Math { .. })));
}
```

- [ ] **Step 3: Run to verify it fails**

Run: `cd packages/docx/parser && cargo test parses_simple_fraction`
Expected: FAIL (no Math handling yet).

- [ ] **Step 4: Implement OMML parsing**

In `parser.rs`, where runs inside a `w:p` are collected, detect `m:oMath` (inline) and `m:oMathPara` (block, sets `display: true`). Walk children recursively:

```rust
fn parse_omath_nodes(el: &Element) -> Vec<MathNode> {
    let mut out = Vec::new();
    for child in el.children() {
        match child.local_name() {
            "r"   => { if let Some(t) = run_text(child) { out.push(MathNode::Run { text: t, style: run_math_style(child) }); } }
            "f"   => out.push(MathNode::Fraction { num: child_nodes(child, "num"), den: child_nodes(child, "den") }),
            "sSup"=> out.push(MathNode::Sup { base: child_nodes(child, "e"), sup: child_nodes(child, "sup") }),
            "sSub"=> out.push(MathNode::Sub { base: child_nodes(child, "e"), sub: child_nodes(child, "sub") }),
            "sSubSup" => out.push(MathNode::SubSup { base: child_nodes(child, "e"), sup: child_nodes(child, "sup"), sub: child_nodes(child, "sub") }),
            "nary"=> out.push(MathNode::Nary { op: nary_char(child), sub: child_nodes(child, "sub"), sup: child_nodes(child, "sup"), body: child_nodes(child, "e") }),
            "d"   => out.push(MathNode::Delimiter { beg_char: delim_char(child, "begChr", "("), end_char: delim_char(child, "endChr", ")"), items: delim_items(child) }),
            "func"=> out.push(MathNode::Func { name: child_nodes(child, "fName"), arg: child_nodes(child, "e") }),
            _ => { let inner = parse_omath_nodes(child); out.extend(inner); }
        }
    }
    out
}
```

Helper `run_math_style` maps `m:rPr/m:sty` (`p`=roman, `i`=italic default, `b`=bold, `bi`) to the TS `MathStyle` strings. `child_nodes(el, name)` finds the `m:<name>` child and parses its `m:e`/content recursively.

- [ ] **Step 5: Run to verify it passes**

Run: `cd packages/docx/parser && cargo test parses_simple_fraction`
Expected: PASS.

- [ ] **Step 6: Rebuild WASM + commit**

```bash
cd packages/docx/parser && wasm-pack build --target web && cp pkg/docx_parser_bg.wasm pkg/docx_parser.js ../src/wasm/
cd ../../.. && git add packages/docx/parser packages/docx/src/wasm
git commit -m "feat(docx): extract OMML equations into shared math AST"
```

---

## Task 8: docx renderer — measure + draw inline math

**Files:**
- Modify: `packages/docx/src/types.ts` (mirror `DocRun::Math`)
- Modify: `packages/docx/src/renderer.ts` (math segment in layout + draw)
- Test: `packages/docx/tests/visual/*` (VRT — local only, refs created with sign-off)

- [ ] **Step 1: Mirror the run type in TS**

In `packages/docx/src/types.ts`, add to the `DocRun` union:

```ts
| { kind: 'math'; nodes: MathNode[]; display: boolean }
```

Import `MathNode` from `@silurus/ooxml-core`.

- [ ] **Step 2: Lazy-load the math font once**

Add a module-level loader in `renderer.ts` (browser: `FontFace` from the bundled asset URL; node: caller-registered). Parse the ArrayBuffer once into `MathFont` + `MathConstants` and cache. Skip entirely if no run has `kind === 'math'`.

```ts
import { parseMathFont, parseMathConstants, layoutMath, renderMathBox, type MathFont, type MathConstants } from '@silurus/ooxml-core';

let mathFontCache: { font: MathFont; consts: MathConstants } | null = null;
async function ensureMathFont(fontUrl: string): Promise<{ font: MathFont; consts: MathConstants }> {
  if (mathFontCache) return mathFontCache;
  const buf = await (await fetch(fontUrl)).arrayBuffer();
  if (typeof FontFace !== 'undefined' && typeof document !== 'undefined') {
    const face = new FontFace('LatinModernMath', buf);
    await face.load();
    (document as Document).fonts.add(face);
  }
  const font = parseMathFont(buf);
  mathFontCache = { font, consts: parseMathConstants(font) };
  return mathFontCache;
}
```

- [ ] **Step 3: Measure math during line layout**

Where segments are measured, for a math run build the box and use its metrics:

```ts
const lctx = { font: mathFontCache!.font, consts: mathFontCache!.consts, fontSizePx: seg.fontSize * scale, level: seg.display ? 'display' : 'text' as const };
const box = layoutMath(seg.nodes, lctx);
seg.measuredWidth = box.width;
// contribute box.ascent / box.descent to line ascent/descent
```

- [ ] **Step 4: Draw math in the segment loop**

In the draw loop (`renderer.ts:1009+`), add a branch before the `LayoutTextSeg` cast:

```ts
if ('nodes' in seg) {  // math segment
  if (!dryRun) {
    const lctx = { font: mathFontCache!.font, consts: mathFontCache!.consts, fontSizePx: seg.fontSize * scale, level: seg.display ? 'display' : 'text' as const };
    const box = layoutMath(seg.nodes, lctx);
    renderMathBox(ctx, box, x, baseline, seg.color ? `#${seg.color}` : defaultColor, 'LatinModernMath');
  }
  x += seg.measuredWidth;
  continue;
}
```

- [ ] **Step 5: Visual smoke test**

Run Storybook, load `private/sample-6.docx` via the Samples story, and confirm equations render (fractions stacked, superscripts raised, sums present). Capture before/after with Claude Preview.

```bash
pnpm build:wasm && pnpm storybook
```

Expected: equations that were previously blank now render as crisp vector math.

- [ ] **Step 6: Typecheck + commit**

```bash
pnpm --filter @silurus/ooxml-docx typecheck
git add packages/docx/src/types.ts packages/docx/src/renderer.ts
git commit -m "feat(docx): render inline OMML equations via core math engine"
```

- [ ] **Step 7: VRT references (ONLY with explicit user sign-off)**

Per CLAUDE.md, reference images are never auto-updated. After visual confirmation, ask the user before running `UPDATE_REFS=1 pnpm --filter @silurus/ooxml-docx vrt` to baseline math.

---

## Self-Review

- **Spec coverage:** AST (Task 1) ✓; font + MATH table (Tasks 2–4) ✓ map to spec "math-table.ts/font.ts"; layout (Task 5) and render (Task 6) ✓; Rust extraction (Task 7) and docx integration (Task 8) ✓. Phase 2/3 items (stretchy delimiters, accents, matrices, pptx/xlsx) are explicitly out of Phase 1 scope per the spec phasing — separate plans.
- **Placeholder scan:** layout/MATH-offset code carries explicit "iterate during TDD" notes with the authoritative spec reference, not vague TODOs; tests assert concrete relationships. Acceptable for research-heavy binary parsing.
- **Type consistency:** `MathNode`/`MathStyle` shared name across core TS and Rust serde tags (`run`/`fraction`/`sup`/`sub`/`subSup`/`nary`/`delimiter`/`func`/`group`); `MathLayoutCtx`/`MathBox`/`DrawOp`/`renderMathBox`/`layoutMath`/`parseMathFont`/`parseMathConstants` consistent across Tasks 5–8.

## Known Risks (carried forward)

1. **MATH constants byte offsets (Task 4):** highest-risk; TDD against a dumped reference until exact.
2. **fillText vs. parsed-advance drift:** Phase 1 measures via parsed hmtx but draws via fillText; advances are the same metric so positions agree. Verify in the smoke test.
3. **Stretchy glyphs deferred:** large radicals/delimiters render at base size in Phase 1 (acceptable, flagged), proper variants in Phase 2 with CFF outlines.
