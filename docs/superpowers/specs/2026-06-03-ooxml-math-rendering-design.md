# OOXML Math (OMML) Vector Rendering — Design

Date: 2026-06-03
Status: Draft (approved direction; spec under review)
Scope: New shared math typesetting + Canvas rendering subsystem in `packages/core`

## Problem

OOXML documents (docx/xlsx/pptx) embed mathematical formulas as **OMML** (Office Math
Markup Language, the `m:` namespace, e.g. `<m:oMath>` / `<m:oMathPara>`). Today none of the
parsers extract this content and none of the renderers draw it — equations are silently
dropped (confirmed: `packages/docx/parser/src/types.rs` has no math types; `renderer.ts`
has no math path). The driving sample is `private/sample-6.docx`, which contains many
equations alongside other unsupported elements (TOC dot leaders, tables).

Math is not docx-specific: the same OMML appears in pptx text bodies and xlsx drawings.
Therefore the solution must be a **single shared mechanism** consumable by all three
viewers, not a per-format feature.

## Constraints (these drive every decision)

1. **Shared across docx/xlsx/pptx.** Lives in `packages/core` (the existing home of
   cross-format Canvas primitives: chart, shape, sparkline).
2. **Canvas vector, no DOM.** The render layer must run against a bare Canvas2D context.
   Browser rendering happens on the main thread with `HTMLCanvasElement`; the `packages/node`
   path uses `skia-canvas` / `@napi-rs/canvas` with **no DOM at all**. Anything depending on
   `document`, HTML elements, `foreignObject`, or html2canvas is disqualified.
   - This is why **KaTeX and MathJax cannot be used as runtime renderers**: KaTeX emits
     HTML+CSS (DOM), MathJax emits SVG/DOM. Neither draws to a Canvas2D context without a DOM.
3. **Crisp at any zoom.** The existing renderer is all-vector (`fillText` + paths); math must
   match. No rasterize-to-bitmap approach (would blur on zoom and diverge across environments).
4. **Beautiful.** Typographic quality is a primary goal; effort is explicitly not a constraint.
5. **Small base bundle.** The dominant cost is the math *font*, not code. Font + engine load
   **lazily, only when a document actually contains math.** Math-free documents see ~zero
   bundle increase.
6. **Spec-faithful over heuristic** (repo CLAUDE.md). Layout constants come from the font's
   OpenType MATH table and ECMA-376 §22 semantics, not empirical magic numbers.

## Chosen Approach

Build a self-contained TeX-style math layout engine that draws vector glyphs sourced from an
OpenType math font, reading layout constants and variant/assembly glyphs from the font's
**OpenType MATH table**. We borrow KaTeX's *layout algorithm and aesthetic* (the source of its
beauty) but not its DOM output or its bundled font set.

### Why not the alternatives

- **KaTeX/MathJax as a library (rasterize or DOM overlay):** violates constraints 2 and 3.
  KaTeX is the worst fit — its output is DOM-bound HTML/CSS and it cannot reach a Canvas2D
  context without a DOM, so it fails the `node` path entirely.
- **Bundle KaTeX's font set + port layout:** the font set is the bulk of the size; a single
  OpenType math font with a real MATH table is smaller and gives us the variant/assembly data
  the MATH table encodes (which KaTeX hardcodes per-font).

## Architecture

```
packages/core/src/types/math.ts   OMML AST — the contract layer (TS types)
packages/core/src/math/
  math-table.ts   minimal OpenType MATH table parser (MathConstants,
                  MathGlyphInfo italic corrections, MathVariants, GlyphAssembly)
  font.ts         lazy math-font loading + glyph metrics + outline (Path) access
  layout.ts       OMML AST -> positioned box tree (TeX mlist->hlist rules)
  render.ts       box tree -> Canvas2D (fillText / Path2D / fillRect)
  index.ts        public surface: measureMath(ast, opts), renderMath(ctx, ast, opts)
```

### Data flow

1. **Parse (Rust, per format).** Each parser extracts `<m:oMath>` / `<m:oMathPara>` into a
   **shared OMML AST** emitted as JSON. The AST shape is defined once as TS types in
   `core/src/types/math.ts`; each Rust parser produces matching JSON. OMML is identical across
   formats, so only the *extraction site* differs per format, not the model.
2. **Layout (`layout.ts`).** AST → box tree of positioned boxes (`x`, `y`, `width`, `height`,
   `baseline`, draw ops). Implements TeX-style math layout: atom classification and inter-atom
   spacing, fraction bar placement, sub/superscript shifts and kerning, radicals, n-ary
   operators with limits, delimiters, accents, matrices. Numeric constants (axis height,
   fraction rule thickness, sub/sup shift-downs, etc.) are read from the MATH table, not
   hardcoded.
3. **Render (`render.ts`).** Walk the box tree and emit Canvas2D calls: cmap-addressable
   glyphs via `fillText`; fraction bars and radical rules via `fillRect`/`Path2D`;
   variant/assembly glyphs (stretchy/large radicals, braces, integrals — frequently have no
   Unicode codepoint) via `Path2D` outlines extracted from the font.

### OMML AST (initial element coverage)

Node kinds mirror OMML elements (ECMA-376 §22.1.2):
`run` (`m:r`), `fraction` (`m:f`), `superscript`/`subscript`/`subSup`
(`m:sSup`/`m:sSub`/`m:sSubSup`), `radical` (`m:rad`), `nary` (`m:nary`), `delimiter` (`m:d`),
`function` (`m:func`), `accent` (`m:acc`), `bar` (`m:bar`), `groupChar` (`m:groupChr`),
`limitLower`/`limitUpper` (`m:limLow`/`m:limUpp`), `matrix` (`m:m`), `box`/`borderBox`
(`m:box`/`m:borderBox`), `phantom` (`m:phant`). Run properties carry the math style
(`m:sty`: roman/italic/bold), scripts, and text. Unknown nodes degrade to rendering their
inner runs rather than disappearing.

## Font Strategy

- **Default math font: Latin Modern Math** (SIL OFL, ~390KB OTF/CFF, Computer Modern aesthetic
  — the look users associate with KaTeX/LaTeX, full OpenType MATH table). Chosen over STIX Two
  Math for smaller size while keeping a complete MATH table.
- **Lazy-loaded** via the existing core font-loading machinery, fetched only when a document
  contains math. Base bundle is unaffected for math-free documents.
- **Configurable** through `LoadOptions` (`core/src/types/load-options.ts`): callers who can
  supply Cambria Math bytes (Word's default math font) may switch to a Word-faithful
  appearance. Faithfulness to Word is therefore an opt-in upgrade, not the default — partly
  because system-installed Cambria Math is **not byte-readable from the browser**, so its MATH
  table cannot be parsed there.
- **Outline access.** Because we own the font bytes, glyphs can be drawn either via `fillText`
  (cmap-addressable) or via `Path2D` outlines (variant/assembly glyphs). Outline extraction +
  the MATH table are read from the same `ArrayBuffer`. Outline rendering guarantees identical
  output across browser main-thread, OffscreenCanvas, and skia-canvas, independent of host font
  rasterizers.

## Integration With Existing Renderers

- `measureMath(ast, opts) -> { width, height, ascent, descent }` lets each format's line
  layout reserve space for inline math, exactly as it already measures text runs.
- `renderMath(ctx, ast, x, baseline, opts)` draws at a baseline-anchored origin, matching the
  `fillText(text, x, baseline)` convention in `packages/docx/src/renderer.ts`.
- docx: inline `m:oMath` inside runs participates in line breaking; block `m:oMathPara` becomes
  its own line/paragraph. pptx/xlsx call the same two functions from their text-body paths.

## Phasing

- **Phase 1 — core engine + docx inline (covers most of sample-6).**
  AST types; MATH-table constant loading; lazy font load; layout + render for fractions,
  sub/superscripts, n-ary operators, ordinary symbols / Greek, and correct inter-atom spacing.
  docx parser extracts `m:oMath`/`m:oMathPara`; docx renderer measures and draws it.
- **Phase 2 — typographic completeness.**
  Variant/assembly stretchy & large radicals and delimiters (MathVariants/GlyphAssembly via
  `Path2D`), accents, over/under bars, `limLow`/`limUpp`, matrices.
- **Phase 3 — pptx/xlsx extraction.**
  Add OMML extraction to the pptx and xlsx Rust parsers; wire their text-body rendering to the
  shared engine. No engine changes expected.

## Testing

- **Unit (TS):** layout produces expected box geometry for representative formulas (fraction
  baseline on the math axis, sub/sup shift relationships, radical rule thickness from the MATH
  table). MATH-table parser decodes known constants from the bundled font.
- **VRT (local only, per repo policy):** new references generated from our renderer for a small
  set of formulas plus `private/sample-6.docx`. Per CLAUDE.md, reference images are **never**
  auto-updated — only on explicit user instruction. Math is a new surface, so initial baselines
  are created deliberately under `UPDATE_REFS=1` with user sign-off.
- Cross-environment check: identical box geometry from `measureMath` in node and browser.

## Risks / Open Questions

- **CFF outline parsing.** Latin Modern Math is CFF-flavored; extracting `Path2D` from CFF is
  more involved than TrueType `glyf`. Phase 1 leans on `fillText` for cmap glyphs to defer this;
  Phase 2 needs a CFF outline reader (own minimal reader vs. a lazily-loaded helper —
  decided during planning, weighed against the bundle-size goal).
- **OMML → AST fidelity in Rust.** Must not collapse `m:sty`, nary properties, or matrix
  alignment during extraction — dropping these forces heuristics later. Extraction is the
  contract; keep it complete.
- **Word-faithful vs. TeX-beautiful default.** Default is TeX-beautiful (Latin Modern Math).
  Revisit only if VRT-against-Word fidelity for math becomes a stated requirement.
