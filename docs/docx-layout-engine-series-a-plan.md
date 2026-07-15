# DOCX Layout Series A Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make body paragraphs, tables, and page/column flow produce one immutable `DocumentLayout` with no parser-model stamps or legacy paint route.

**Architecture:** Introduce stable result and service contracts first, then migrate paragraphs, in-flow/nested tables, floating/split tables, and finally the page state machine. Every paint module consumes self-contained geometry and resources and cannot import measurement code.

**Tech Stack:** TypeScript, Canvas 2D, Vitest, ast-grep, pnpm, Rust-generated DOCX models.

## Global Constraints

- Follow all constraints and the per-PR independent review gate in `docx-layout-engine-implementation-roadmap.md`.
- Preserve public API signatures and render behavior.
- Do not retain a migrated feature's old algorithm behind a flag or predicate.
- Layout code cannot read scale, DPR, Canvas state, or paint callbacks.
- Paint code cannot measure, shape, paginate, resolve styles, or dereference parser objects.

---

### Task A1: Establish immutable layout, diagnostics, invariants, and paint purity

**Files:**

- Create: `packages/docx/src/layout/types.ts`
- Create: `packages/docx/src/layout/diagnostics.ts`
- Create: `packages/docx/src/layout/compatibility.ts`
- Create: `packages/docx/src/layout/compatibility.test.ts`
- Create: `packages/docx/src/layout/invariants.ts`
- Create: `packages/docx/src/layout/flow.ts`
- Create: `packages/docx/src/layout/invariants.test.ts`
- Create: `packages/docx/src/layout/structured-clone.test.ts`
- Create: `packages/docx/src/paint/canvas-page.ts`
- Create: `packages/docx/src/paint/canvas-drawing.ts`
- Create: `packages/docx/src/paint/paint-purity.test.ts`
- Create: `docs/docx-layout-shared-primitives-audit.md`
- Read for the audit: `packages/core/src/types/common.ts`
- Read for the audit: `packages/core/src/fonts/font-registry.ts`
- Read for the audit: `packages/core/src/fonts/local-metrics.ts`
- Read for the audit: `packages/core/src/image/raster-dimensions.ts`
- Read for the audit: `packages/core/src/chart/renderer.ts`
- Read for the audit: `packages/ooxml-common/src/drawing.rs`
- Read for the audit: `packages/ooxml-common/src/text.rs`
- Read for the audit: `packages/ooxml-common/src/color.rs`
- Read for the audit: `packages/ooxml-common/src/units.rs`
- Create: `packages/docx/api/public-api-baseline.d.ts`
- Create: `scripts/check-docx-layout-boundaries.mjs`
- Create: `scripts/check-docx-layout-boundaries.test.mjs`
- Create: `scripts/docx-layout-boundary-baseline.json`
- Create: `scripts/check-docx-public-api.mjs`
- Create: `rules/no-docx-layout-in-paint.yml`
- Create: `rules/no-docx-display-scale-in-layout.yml`
- Create: `rules/no-docx-style-resolution-in-layout-paint.yml`
- Create: `rule-tests/no-docx-layout-in-paint-test.yml`
- Create: `rule-tests/no-docx-display-scale-in-layout-test.yml`
- Create: `rule-tests/no-docx-style-resolution-in-layout-paint-test.yml`
- Modify: `sgconfig.yml`
- Modify: `package.json`
- Modify: `.github/workflows/ci.yml`

**Interfaces:**

- Consumes: `SectionLayoutContext` from `packages/docx/src/layout-context.ts` and parser model types from `types.ts` only at the layout boundary.
- Produces: `DocumentLayout`, `LayoutPage`, `PageLayers`, `PaintNode`, `LayoutDiagnostic`, `layoutFlowBlocks`, `assertDocumentLayout`, `layoutFingerprint`, a public declaration baseline, and the transitive dependency checker.

**Specification evidence:** This boundary PR does not implement a new OOXML
layout rule. Its audit must map existing shared DrawingML transforms, colors,
images, charts, and font primitives to ECMA-376 Parts 1 and 4, and record why
WordprocessingML page flow, stories, and table pagination remain DOCX-local.

- [ ] **Step 1: Write failing invariant and paint-purity tests**

Add tests that construct two ordinary in-flow allocation ranges with overlapping
`flowBounds`, an ordinary in-flow node whose allocation crosses
`geometry.contentBottomPt`, allowed floating overlap, negative-spacing ink that
extends beyond its allocation, a clipped frame overhang, a NaN coordinate, and a
Canvas stub whose `measureText()` throws. Assert only ordinary flow ownership
throws `FLOW_OVERLAP` or `BOTTOM_MARGIN_INVASION`; allowed ink/floating overlap
passes; NaN throws `INVALID_GEOMETRY`; minimal paint succeeds without measuring.

```ts
expect(() => assertDocumentLayout(overlappingLayout)).toThrow(/FLOW_OVERLAP/);
expect(() => assertDocumentLayout(marginLayout)).toThrow(/BOTTOM_MARGIN_INVASION/);
expect(() => assertDocumentLayout(nanLayout)).toThrow(/INVALID_GEOMETRY/);
expect(() => assertDocumentLayout(allowedFloatOverlap)).not.toThrow();
expect(() => assertDocumentLayout(negativeInkOverhang)).not.toThrow();
expect(() => paintLayoutPage(paintOnlyLayout, 0, canvas, { scale: 1, dpr: 1 }))
  .not.toThrow();
```

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/invariants.test.ts packages/docx/src/paint/paint-purity.test.ts
```

Expected: FAIL because the new modules and exports do not exist.

- [ ] **Step 3: Add the minimal immutable contracts**

Define deep-readonly, structured-clone-safe point-space data and diagnostics:

```ts
export interface LayoutRect { readonly xPt: number; readonly yPt: number; readonly widthPt: number; readonly heightPt: number }
export interface PageGeometry extends LayoutRect { readonly contentTopPt: number; readonly contentBottomPt: number }
export interface FlowOwnership { readonly flowBounds: LayoutRect; readonly inkBounds: LayoutRect; readonly clipBounds?: LayoutRect; readonly advancePt: number; readonly ordinaryFlow: boolean }
export type LayoutDiagnosticCode = 'FLOW_OVERLAP' | 'BOTTOM_MARGIN_INVASION' | 'INVALID_GEOMETRY' | 'NON_CONVERGENCE' | 'UNSUPPORTED_FEATURE';
export interface LayoutDiagnostic { readonly code: LayoutDiagnosticCode; readonly severity: 'warning' | 'error'; readonly source?: SourceRef; readonly message: string }
export type CompatibilityEvidence = Readonly<{ kind: 'microsoft-note'; reference: string }> | Readonly<{ kind: 'office-observation'; syntheticFixtureId: string; application: string; version: string; platform: string }>;
export interface CompatibilityRule { readonly id: string; readonly evidence: CompatibilityEvidence; readonly description: string }
export function assertDocumentLayout(layout: DocumentLayout): void;
export function layoutFingerprint(layout: DocumentLayout): string;
export interface BlockLayoutAlgorithms {
  layoutParagraph(input: ParagraphLayoutInput, services: LayoutServices): ParagraphLayout;
  layoutTable(input: TableLayoutInput, services: LayoutServices): TableLayout;
}
export type FlowBlockInput = ParagraphLayoutInput | TableLayoutInput;
export interface FlowLayoutInput { readonly blocks: readonly FlowBlockInput[]; readonly container: FlowContainer; readonly source: SourceRef }
export function layoutFlowBlocks(input: FlowLayoutInput, services: LayoutServices, algorithms: BlockLayoutAlgorithms): FlowLayout;
```

Implement `layoutFingerprint` by recursively normalizing finite numbers to six
decimal places and serializing pages, layers, and reading order; omit diagnostics'
free-form message but include code, severity, and source.

Define transforms as numeric `Matrix2DData`, clips as plain-data discriminated
unions, and `SourceRef.storyInstance` as `body`, a header/footer part key, a note
kind plus note ID, or the enclosing shape source path. Add a construction-time
`deepFreezeDocumentLayout` and assert `structuredClone(layout)` preserves the
normalized result without DOM/Canvas/function/WeakMap values.

Add tested directory rules alongside the transitional single-file rule; A3
deletes that old rule and its test with `fragment-paint.ts`. Paint may use
`import type` from `layout/types.ts`, but value imports from layout
algorithm modules and calls/imports of `measureText`, `layoutLines`,
`computeTableLayout`, `paginateDocument`, shaping, or style-resolution functions
are rejected. Layout rejects Canvas context types, `PaintPageOptions`,
`CanvasPaintContext`, `.dpr`, and `.displayScale`; it does not reject OOXML
authored horizontal `w:w` scale fields by spelling alone.

The third rule rejects style-cascade/property-merge helper declarations and
imports in `layout/**` and `paint/**`; authored/inherited/cleared/absent style
resolution remains in the Rust parser and the typed parser boundary. Layout
context construction may map resolved facts but cannot re-run the cascade.

`scripts/check-docx-layout-boundaries.mjs` follows the complete local import graph
from every `paint/*.ts` entry and rejects new transitive reachability to
measurement, pagination, style-resolution, or parser-object modules. In A1,
`--write-transitional-baseline` records the exact existing `renderer.ts`
layout/measurement declarations and import edges in
`scripts/docx-layout-boundary-baseline.json`; ordinary checks fail on additions.
Every later migration PR removes the entries it deletes and may never add an
allowance. C3 deletes the empty baseline and runs `--final`, which requires the
adapter-only export/import allowlist and zero transitive forbidden edges; ordinary
checks also enter final mode automatically once the baseline file is absent. Generate and commit the
current `@silurus/ooxml/docx` declaration surface as
`packages/docx/api/public-api-baseline.d.ts` before production migration begins.
`scripts/check-docx-public-api.mjs --write-baseline` writes the normalized
baseline only in A1; later invocations omit that flag and fail on any declaration
change.

For every post-A1 branch, the boundary checker computes `git merge-base
origin/main HEAD`, reads that commit's baseline with `git show`, and enforces
`headAllowances ⊆ mergeBaseAllowances` before checking source edges. A changed
JSON file therefore cannot authorize a new edge. The A1-only
`--write-transitional-baseline` command is valid only when the merge base has no
baseline file; it fails once a baseline exists. `package.json` exposes
`test:docx-boundaries` as `node --test
scripts/check-docx-layout-boundaries.test.mjs`, and CI runs it with lint. Set the
CI `actions/checkout` step to `fetch-depth: 0`, so `origin/main`, the merge base,
and its committed baseline are available to the checker rather than depending on
a shallow checkout's incidental history.

- [ ] **Step 4: Run focused and static checks**

Run:

```bash
pnpm vitest run packages/docx/src/layout/invariants.test.ts packages/docx/src/paint/paint-purity.test.ts
pnpm lint
pnpm lint:test
pnpm test:docx-boundaries
node scripts/check-docx-layout-boundaries.mjs --write-transitional-baseline
node scripts/check-docx-layout-boundaries.mjs
pnpm build
node scripts/check-docx-public-api.mjs --write-baseline
pnpm typecheck
```

Expected: all commands pass; rule tests prove type-only node imports are valid
and algorithm/measurement/display-scale access fails. Node negative fixtures
prove a new forbidden import edge fails, expanding the head baseline beyond the
merge-base set fails, and `--final` with a nonempty baseline fails. The
structured-clone test contains no live platform objects.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A1**

Commit subject: `refactor(docx): establish immutable layout boundaries`.
Use the roadmap review gate and merge only with all checks and findings clear.

### Task A2: Establish stable layout services, options, resources, and convergence

**Files:**

- Create: `packages/docx/src/layout/font-service.ts`
- Create: `packages/docx/src/layout/font-service.test.ts`
- Create: `packages/docx/src/layout/text.ts`
- Create: `packages/docx/src/layout/resources.ts`
- Create: `packages/docx/src/layout/resources.test.ts`
- Create: `packages/docx/src/layout/convergence.ts`
- Create: `packages/docx/src/layout/convergence.test.ts`
- Create: `packages/docx/src/layout/options.ts`
- Create: `packages/docx/src/layout/options.test.ts`
- Create: `packages/docx/src/layout/error-page.ts`
- Create: `packages/docx/src/layout/error-page.test.ts`
- Create: `packages/docx/src/layout/numbering-marker.ts`
- Create: `packages/docx/src/paint/numbering-marker.ts`
- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Modify: `packages/docx/src/parser-model.ts`
- Modify: `packages/docx/src/line-layout.ts`
- Modify: `packages/docx/src/paragraph-measure.ts`
- Modify: `packages/docx/src/local-font-metrics.ts`
- Modify: `packages/docx/src/embedded-fonts.ts`
- Modify: `packages/docx/src/google-fonts.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/document.ts`
- Modify: `packages/docx/src/render-worker.ts`
- Modify: `packages/docx/src/worker-protocol.ts`
- Modify: `scripts/check-docx-layout-boundaries.mjs`
- Modify: `scripts/check-docx-layout-boundaries.test.mjs`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: A1 deep-readonly/plain-data types,
  `packages/core/src/fonts/font-registry.ts`,
  `packages/core/src/fonts/local-metrics.ts`,
  `packages/core/src/image/raster-dimensions.ts`, and existing shared chart/math
  paint primitives as classified by `docs/docx-layout-shared-primitives-audit.md`.
- Produces: the final `TextLayoutService`, `ImageMetadataService`, `MathMetadataService`, `FontResolver`, `LayoutOptions`, `layoutOptionsKey`, `convergeLayout`, and parse-error `DocumentLayout` contracts used unchanged by all later PRs.

```ts
export interface CanvasFontRoute { readonly familyList: string; readonly scope: 'registered' | 'native' | 'generic'; readonly fingerprint: string }
export interface FontResolution { readonly requestedFamily: string; readonly resolvedFamily: string; readonly source: 'embedded' | 'local' | 'google' | 'substitute' | 'native' | 'generic'; readonly route: CanvasFontRoute; readonly weight: number; readonly style: 'normal' | 'italic'; readonly diagnostics: readonly LayoutDiagnostic[] }
export interface FontResolver { resolve(request: Readonly<FontRequest>): FontResolution }
export interface TextLayoutService { readonly fingerprint: string; resolve(request: Readonly<TextFontResolveRequest>): FontResolution; shape(request: Readonly<TextShapeRequest>): TextShapeResult }
export interface ImageMetadataService { readonly fingerprint: string; resolve(resourceKey: string): Readonly<{ widthPt: number; heightPt: number; mimeType: string }> }
export interface MathMetadataService { readonly fingerprint: string; resolve(resourceKey: string): DeepReadonly<MathLayoutResource> }
export interface NumberingMarkerShapeInput { readonly fontSizePt: number; readonly fonts: TextFontSlots; readonly themeFonts?: TextFontSlots; readonly themeFontPresence?: TextFontSlotPresence; readonly weight: number; readonly style: 'normal' | 'italic'; readonly complexScript: boolean; readonly fontHint?: 'default' | 'eastAsia' | 'cs'; readonly eastAsiaLanguage?: string; readonly kerning?: boolean }
export interface LayoutOptions { readonly currentDateMs: number }
export function layoutOptionsKey(options: LayoutOptions, services: LayoutServices): string;
export function convergeLayout(seed: LayoutIteration, step: (iteration: LayoutIteration) => LayoutIteration, limit: number): LayoutIteration;
```

`CanvasFontRoute` is a format-neutral, immutable CSS request created and
serialized by core. DOCX retains Word-specific four-slot/theme/fontTable policy.
An uninventoried authored family uses an engine-scoped `native` route with an
explicit metadata-derived generic tail; this makes no portable availability or
geometry claim. Inventory faces are exact `(family, weight, style)` tuples, not
Cartesian products. Service fingerprints identify immutable resources and route
syntax; A3's retained/document-layout geometry is the downstream boundary. It
preserves same-browser main/worker parity without claiming cross-browser Canvas
geometry portability; cross-browser guarantees remain semantic invariants.

The parser wire may retain private effective numbering-level and paragraph-mark
font facts, but `parser-model.ts` is their projection boundary. It snapshots
those facts into immutable plain inputs such as `NumberingMarkerShapeInput`.
The layout numbering-marker module consumes only that input and the text service,
and the paint module serializes only the retained shaped spans. Neither layout
nor paint dereferences the private parser wire representation.

**Specification evidence:** ECMA-376 §17.3.2.26 (`w:rFonts`), §17.8 embedded
fonts, §17.16.5.13/§17.16.5.65 DATE/TIME, §17.16.5.42 NUMPAGES and
§17.16.5.44 PAGE define layout-affecting font/field facts. Numbering-level
properties apply to the marker under §17.9.6 (`w:lvl`) together with the
§17.3.2.26 font-slot rules. Paragraph-mark run properties are defined by
§17.3.1.29 (`w:pPr`). Table auto-fit selection is defined by §17.4.52
(`w:tblLayout`). DrawingML inline and
anchor extents (`wp:extent`) supply image/chart intrinsic layout size. Font
substitution is environment/Office compatibility behavior and must emit a
resolution record, not hide inside paragraph geometry.

- [ ] **Step 1: Write failing service, option, convergence, and error-page tests**

Use fake font inventories and glyph measurers to cover ASCII, East Asian,
complex-script, theme, embedded, local, Google, missing, bold, and italic
resolution. Use fake image/math resources to assert stable string resource keys
and plain metadata. Assert `layoutOptionsKey` changes for `currentDateMs` or any
service-owned resource fingerprint but not paint width/DPR/default color. Prove
there is no overload accepting caller-supplied environment strings. Assert convergence returns on
a stable fingerprint, throws `NON_CONVERGENCE` on a repeated cycle or limit, and
never returns a stale iteration. Assert parse-error text wraps during layout,
retains the exact `CanvasFontRoute` used for that measurement, and paints with
the same core serialization without calling `measureText`.

Add focused Red cases proving that auto-fit preserves the requested `hAnsi`,
theme, and substitute route on every grapheme/kinsoku segment and sums
differently formatted `joinPrev` pieces per route. Cover body and text-box
numbering for mixed `第1章` and `U+2022` markers across all four font slots and
theme precedence, retaining the shaped spans for paint. Cover empty and
anchor-only paragraph marks in main and worker mode through the same text-service
metrics. Boundary tests must reject direct and transitive parser-model reachability
through bridge modules, aliases, literal/non-literal dynamic imports, and
CommonJS `require`, while proving that only the exact unaliased named runtime
import `{ normalizeInternalDocumentModel }` from `../parser-model.js` in
`layout/resources.ts` is terminal, and that binding is used only by the exact
`documentMathOccurrences` projection to return `.mathOccurrences`. Local
re-export, alias, leak, or extra reference cases fail. All other gateway edges
remain traversed; erased type-only contracts remain valid.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/font-service.test.ts packages/docx/src/layout/resources.test.ts packages/docx/src/layout/options.test.ts packages/docx/src/layout/convergence.test.ts packages/docx/src/layout/error-page.test.ts
```

Expected: FAIL because services/options/convergence do not exist and
`drawParseErrorPlaceholder` wraps with Canvas `measureText` during paint.

- [ ] **Step 3: Implement final instance-scoped services before feature migration**

Snapshot available font families and resources into immutable per-document
service instances, reuse format-neutral core resource primitives, and keep DOCX
theme/script selection local. Retrofit the existing renderer to consume this
single service instance so no temporary second font/shaping implementation is
introduced. Capture the load-time default `Date.now()` once on the main thread,
send `defaultCurrentDateMs` as an internal worker parse field, and normalize every
`Date | number | undefined` to `LayoutOptions.currentDateMs`. Derive
`layoutOptionsKey(options, services)` from that option and the three actual
service fingerprints; A6 uses it when `DocumentLayout` becomes the retained
cache value. Implement convergence with a relevant geometry fingerprint plus a
seen-set and hard error. Convert the parse-error placeholder into stored
text/frame paint nodes.

Delete module-global document-specific resolved-font state and paint-time parse
error wrapping. Replace parser-object-keyed math `WeakMap` state with stable
`SourceRef`/resource-key records owned by `MathMetadataService`; the existing
worker-safe fallback is retained as an explicit diagnostic result when the
optional DOM math engine is unavailable. Keep glyph shaping in the injected
service; later PRs consume it without replacing its algorithm or interface.

Route auto-fit through `buildSegments` and preserve the complete service request
per atom so region widths sum the actual shaped advances even across font-route
and `joinPrev` boundaries. Route empty and anchor-only paragraph-mark metrics
through that same service authority. Preserve effective numbering and
paragraph-mark run facts on the private parser wire, project numbering at the
renderer boundary into `NumberingMarkerShapeInput`, and retain the service-shaped
body/text-box marker spans for paint. This implements the whole class defined by
the cited `w:rFonts`, numbering-level, and paragraph-mark rules rather than
tuning one marker string.

Keep the transitional `renderShapeText` hash normalization mechanically exact:
it may erase only the complete marker snapshot -> service shape -> retained-span
paint sequence, and must reject any partial or altered sequence. Enforce the
layout/parser boundary with a transitive runtime import-graph walk from every
production layout module. Reject indirect bridge, alias, dynamic-import, and
CommonJS paths. Treat only the exact unaliased named runtime import
`{ normalizeInternalDocumentModel }` from `../parser-model.js` in
`layout/resources.ts` as a terminal parser projection edge, and AST-freeze its
sole use to the exported `documentMathOccurrences` return of
`[...normalizeInternalDocumentModel(doc).mathOccurrences]`. Reject local
re-export, alias, leak, or extra references; traverse every other gateway edge
normally. Erased type-only contracts do not create runtime paths.

- [ ] **Step 4: Verify Green and main/worker service parity**

Run:

```bash
pnpm vitest run packages/docx/src/layout/{font-service,resources,options,convergence,error-page}.test.ts
pnpm vitest run packages/docx/src/layout/services-integration.test.ts packages/docx/src/fit-text-fixes.test.ts packages/docx/src/column-widths.test.ts packages/docx/src/paragraph-measure.test.ts packages/docx/src/empty-paragraph-mark-height.test.ts packages/docx/src/numbering-marker-font.test.ts
cargo test -p docx-parser
pnpm test:docx-boundaries
node scripts/check-docx-layout-boundaries.mjs --base-ref 02863444
pnpm test:docx-public-api
pnpm test:docx-public-api -- --exact
rg -n 'drawParseErrorPlaceholder|setResolvedLocalFonts|clearResolvedLocalFonts' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: tests pass; main and worker factories given identical inventories
produce identical resolution/service fingerprints; auto-fit, numbering, and
paragraph marks use that service in both modes; direct and transitive parser-model
paths fail except for the exact normalization import, while all other gateway
edges are inspected and type-only contracts stay erased; the exact public API
surface is unchanged; and `rg` has no production global-state or
paint-error-wrapper matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A2**

Commit subject: `refactor(docx): establish deterministic layout services`.
Use the roadmap review gate.

### Task A3: Route every body paragraph and run resource through self-contained layout

**Files:**

- Modify: `packages/docx/src/layout/text.ts`
- Create: `packages/docx/src/layout/paragraph.ts`
- Create: `packages/docx/src/layout/paragraph.test.ts`
- Create: `packages/docx/src/layout/run-resources.test.ts`
- Create: `packages/docx/src/layout/textbox-input.ts`
- Create: `packages/docx/src/layout/textbox-input.test.ts`
- Create: `packages/docx/src/paint/canvas-text.ts`
- Create: `packages/docx/src/paint/canvas-text.test.ts`
- Modify: `packages/docx/src/paint/canvas-drawing.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/layout-fragments.ts`
- Delete: `packages/docx/src/fragment-paint.ts`
- Delete: `rules/no-docx-measurement-in-fragment-paint.yml`
- Delete: `rule-tests/no-docx-measurement-in-fragment-paint-test.yml`
- Modify: `packages/docx/src/fragment-paint.test.ts`
- Modify: `packages/docx/src/layout-lines-reuse-identity.test.ts`
- Modify: `packages/docx/src/layout-lines-scale-invariance.test.ts`
- Modify: `packages/docx/src/layout-lines-zoom-invariant.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: the stable services, options, convergence primitive, `SourceRef`, and invariant contracts from A1/A2.
- Produces: `ParagraphLayout`, `TextPlacement`, `InlineResourceLayout`, `DrawingLayout`, `TextBoxLayout`, `normalizeTextBoxInput`, `layoutParagraph`, and `paintParagraphLayout`.

**Specification evidence:** ECMA-376 §17.3.1.13 (`w:jc`), §17.3.1.38
(`w:tabs`), §17.3.1.33 (`w:spacing`), §17.3.1.19/§17.9 numbering,
§17.3.2.41 (`w:vanish`), §17.16 fields, §17.6.20 text direction,
§20.4.2.8 (`wp:inline`), §20.4.2.7 (`wp:extent`), and §20.4.2.3
(`wp:anchor`). Picture bullets follow
§17.9.9/§17.9.20. DrawingML chart paint remains the shared core implementation;
DOCX owns only inline/anchor flow placement.

```ts
export interface ParagraphLayout {
  readonly kind: 'paragraph';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly advancePt: number;
  readonly lines: readonly LineLayout[];
  readonly borders: readonly BorderSegment[];
  readonly shading?: FillPaint;
}
export function layoutParagraph(input: ParagraphLayoutInput & Readonly<{ exclusions: readonly WrapExclusion[] }>, services: LayoutServices): ParagraphLayout;
export function normalizeTextBoxInput(shape: ShapeRun): readonly ParagraphLayoutInput[];
export function paintParagraphLayout(node: ParagraphLayout, context: CanvasPaintContext): void;
```

- [x] **Step 1: Add failing behavior tests**

Add synthetic tests for numbered paragraphs, bidi runs, vertical text, tab
leaders, floating-wrap exclusions, continuation slices, contextual spacing,
paragraph borders, hidden paragraph marks, and page fields. Assert exact line
text ranges and point bounds, and assert the paginator's cursor delta equals
`advancePt` while ink/clip height may differ from it. Two paints at scale 1 and 2
must not call `measureText` or change the layout fingerprint.

Add one matrix-driven test covering every `DocRun` arm and associated resource:

| Input | Required retained node/resource | Owning implementation |
|---|---|---|
| `text` | shaped `TextPlacement` | `layout/text.ts` |
| `anchorHost` | anchor baseline/source metrics | `layout/paragraph.ts` |
| inline/anchored `image` | resource key, intrinsic size, bounds/wrap facts | `layout/resources.ts` + `DrawingLayout` |
| inline/anchored `chart` | shared chart resource key and bounds | core chart paint + `DrawingLayout` |
| line/page/column `break` | line or flow event | paragraph/paginator |
| `field` | resolved text plus field dependency | paragraph + A2 convergence |
| `shape` / text box | drawing bounds plus existing public `textBlocks` converted to retained paragraph layouts; B2 generalizes the same normalizer to richer blocks while preserving this fallback | `DrawingLayout` + `TextBoxLayout` |
| `math` | stable math resource key and layout bounds | A2 `MathMetadataService` |
| `ptab` | positioned tab placement | paragraph |
| picture bullet | image resource key and marker bounds | paragraph/resources |

- [x] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/paragraph.test.ts packages/docx/src/paint/canvas-text.test.ts packages/docx/src/layout-lines-reuse-identity.test.ts packages/docx/src/layout-lines-scale-invariance.test.ts packages/docx/src/layout-lines-zoom-invariant.test.ts
```

Expected: new contract tests fail because paint still delegates to
`renderBodyParagraphLines` and scale-2 paint remeasures.

- [x] **Step 3: Move line acquisition and glyph geometry into layout**

Adapt existing `buildSegments`, bidi/tab resolution, `layoutLines`, numbering,
field resolution, and paragraph decoration calculations into `layout/text.ts`
and `layout/paragraph.ts`. Materialize all matrix entries as text, inline
resource, drawing, break, or paginator-event data. Store resolved glyph text,
font descriptor, advances, offsets, decorations, link/bookmark metadata, and
resource keys on `TextPlacement`. `canvas-text.ts` and `canvas-drawing.ts` only
apply stored transforms and call drawing primitives.

For a shape carrying the existing public `ShapeRun.textBlocks`,
`normalizeTextBoxInput` converts each `ShapeText` compatibility block to a
`ParagraphLayoutInput`, lays it out through the same `layoutParagraph`, and
retains those paragraph nodes inside `TextBoxLayout`. This function is the single
permanent input-normalization boundary for both parser-produced and externally
constructed public models. Delete `renderShapeText` measurement and parser
dereference in this PR. B2 extends this normalizer to prefer full internal
`textBoxContent` plus `layoutStory`, while retaining the public `textBlocks`
fallback. It reuses the same paragraph/table nodes and paint contract and
therefore introduces no temporary second text-box algorithm.

Paragraph layout consumes only immutable `WrapExclusion` polygons; it does not
place or retry floats. Until C1, the single existing float placer is adapted to
produce that contract. C1 replaces that provider and deletes its mutable logic
without changing or duplicating paragraph line layout.

Delete `fitMeasureReuseEnabled`, `fragmentPaintEnabled`,
`lineReuseEnabled`, `isFragmentPaintableParagraph`, `layoutLinesInputs`, and
`stampParagraphLines`. Remove `source: DocParagraph` and `MeasuredParagraph` from
paint-facing fragments; retain only `SourceRef` and self-contained paint data.

- [x] **Step 4: Verify Green and prove deletion**

Run:

```bash
pnpm vitest run packages/docx/src/layout/paragraph.test.ts packages/docx/src/layout/run-resources.test.ts packages/docx/src/layout/textbox-input.test.ts packages/docx/src/paint/canvas-text.test.ts packages/docx/src/fragment-paint.test.ts packages/docx/src/layout-lines-reuse-identity.test.ts packages/docx/src/layout-lines-scale-invariance.test.ts packages/docx/src/layout-lines-zoom-invariant.test.ts
rg -n 'fitMeasureReuseEnabled|fragmentPaintEnabled|lineReuseEnabled|isFragmentPaintableParagraph|layoutLinesInputs|stampParagraphLines|renderBodyParagraphLines|renderShapeText' packages/docx/src
pnpm lint
pnpm lint:test
pnpm typecheck
```

Expected: tests and typecheck pass; `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A3**

Commit subject: `refactor(docx): make paragraph paint consume layout geometry`.
Use the roadmap review gate.

### Task A4: Build in-flow and nested table geometry from one measurement

**Files:**

- Create: `packages/docx/src/layout/table.ts`
- Create: `packages/docx/src/layout/table.test.ts`
- Create: `packages/docx/src/paint/canvas-table.ts`
- Create: `packages/docx/src/paint/canvas-table.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/table-fragments.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/table-layout-reuse.test.ts`
- Modify: `packages/docx/src/cell-border-conflict-render.test.ts`
- Modify: `packages/docx/src/column-widths.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: `layoutParagraph` and A1's recursive `layoutFlowBlocks` coordinator.
- Produces: `TableLayout`, `TableRowLayout`, `TableCellLayout`, `ResolvedBorderSegment`, `layoutTable`, and `paintTableLayout`.

**Specification evidence:** ECMA-376 §17.4.37 (`w:tbl`), §17.4.48
(`w:tblGrid`), §17.4.52 (`w:tblLayout`), §17.4.80 (`w:trHeight`),
§17.4.84 (`w:vMerge`), §17.4.68 (`w:tcMar`), §17.4.71 (`w:tcW`), and
the table/cell border conflict rules in §17.4 define the retained geometry.

```ts
export interface TableLayout {
  readonly kind: 'table';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly advancePt: number;
  readonly columnWidthsPt: readonly number[];
  readonly rows: readonly TableRowLayout[];
  readonly borders: readonly ResolvedBorderSegment[];
}
export function layoutTable(input: TableLayoutInput, services: LayoutServices): TableLayout;
export function paintTableLayout(node: TableLayout, context: CanvasPaintContext): void;
```

- [ ] **Step 1: Write failing single-acquisition tests**

Create a counting `TextLayoutService` and synthetic fixed/auto tables containing
paragraphs, nested tables, vertical merges, row spans, exact/at-least heights,
cell margins, and conflicting borders. Assert each paragraph is shaped once per
placement, row heights equal the sum/max of retained child layouts, and paint
does not increment the counter.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table.test.ts packages/docx/src/paint/canvas-table.test.ts packages/docx/src/table-layout-reuse.test.ts packages/docx/src/cell-border-conflict-render.test.ts packages/docx/src/column-widths.test.ts
```

Expected: counting assertions fail because `buildTableCellBlocks` performs a
second cell-content measurement and paint retains a legacy supplied-geometry
bridge.

- [ ] **Step 3: Implement one retained table acquisition**

Resolve the grid, lay out each cell's blocks once, compute intrinsic cell heights
from those retained blocks, resolve row/vMerge heights, translate child bounds to
final cell positions, and resolve shared border segments once. Recursively use
the same function for nested tables. `paintTableLayout` draws stored backgrounds,
children, clipping, and border segments only.

Remove `tableColWidthsPt`, `tableRowHeightsPt`, and `tableLayoutInputs` from
`PaginatedBodyElement`; delete their writes and reuse checks for in-flow and
nested tables. Delete the second paragraph acquisition in
`buildTableCellBlocks`; preserve a single function that converts retained child
layouts into page fragments.

- [ ] **Step 4: Verify Green and mutation safety**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table.test.ts packages/docx/src/paint/canvas-table.test.ts packages/docx/src/table-layout-reuse.test.ts packages/docx/src/cell-border-conflict-render.test.ts packages/docx/src/column-widths.test.ts
rg -n 'tableColWidthsPt|tableRowHeightsPt|tableLayoutInputs' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: tests pass, parser input remains deeply equal before/after layout and
paint, and `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A4**

Commit subject: `refactor(docx): retain one table layout acquisition`.
Use the roadmap review gate.

### Task A5: Migrate floating and page-split tables without a legacy gate

**Files:**

- Modify: `packages/docx/src/layout/table.ts`
- Create: `packages/docx/src/layout/table-pagination.test.ts`
- Modify: `packages/docx/src/table-fragments.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/float-table-geometry.test.ts`
- Modify: `packages/docx/src/float-table-page-fit.test.ts`
- Modify: `packages/docx/src/float-table-width.test.ts`
- Modify: `packages/docx/src/pagination.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: `TableLayout` from A4 and current float exclusion inputs.
- Produces: `TableFragmentLayout`, per-cell continuation ranges, `splitTableLayout`, and floating `TableLayout` placements using the same node type.

**Specification evidence:** ECMA-376 §17.4.6 (`w:cantSplit`), §17.4.49
(`w:tblHeader`), §17.4.84 (`w:vMerge`), and §17.4.57 (`w:tblpPr`) define
row splitting, repeated headers, merged-cell continuation, and floating table
positioning. Office-observed mid-cell pagination behavior must be isolated and
documented if the normative text does not determine a unique split.

```ts
export interface BlockContinuationRange { readonly blockIndex: number; readonly lineStart?: number; readonly lineEnd?: number; readonly childFragmentIndex?: number }
export interface TableCellFragmentLayout { readonly logicalCellIndex: number; readonly contentRanges: readonly BlockContinuationRange[]; readonly flowBounds: LayoutRect }
export interface TableRowFragmentLayout { readonly logicalRowIndex: number; readonly fragmentIndex: number; readonly ownership: 'source' | 'repeated-header'; readonly cells: readonly TableCellFragmentLayout[]; readonly flowBounds: LayoutRect }
export interface TableFragmentLayout { readonly tableId: LayoutNodeId; readonly rows: readonly TableRowFragmentLayout[]; readonly continuesFromPreviousPage: boolean; readonly continuesOnNextPage: boolean }
export function splitTableLayout(table: TableLayout, availableHeightPt: number): readonly TableFragmentLayout[];
```

- [ ] **Step 1: Write failing continuation and floating tests**

Cover repeated header rows, `cantSplit`, mid-cell paragraph continuation,
vertical merge continuation, nested table continuation, negative table indent,
floating table wrapping, and a float that must move to the next page. Include a
`PAGE` field in a repeated header and assert every emitted header occurrence is
acquired with its own physical/display page context. Also place `PAGE` in a
non-leading cell block that continues onto another page and assert its retained
source uses the original `BlockContinuationRange.blockIndex`, not the slice-local
block index. Assert
logical rows may have several disjoint fragments, per-cell content ranges are
disjoint and exhaustive, repeated headers use `repeated-header` ownership and do
not claim source content twice, fragments reuse the same resolved columns, and
ordinary flow bounds do not enter the bottom margin.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table-pagination.test.ts packages/docx/src/float-table-geometry.test.ts packages/docx/src/float-table-page-fit.test.ts packages/docx/src/float-table-width.test.ts packages/docx/src/pagination.test.ts
```

Expected: legacy-gated floating cases do not produce `TableLayout` continuations.

- [ ] **Step 3: Implement splits as immutable views over retained geometry**

Split only at legal row/cell/line boundaries, repeat the resolved header layout,
and create fragment-local translated bounds without recomputing columns or text.
Represent floating placement as a `DrawingLayout`-style placement wrapper whose
child is the same `TableLayout` used for in-flow content.

Delete `tableRequiresLegacyPaint`, `isFragmentPaintableTable`,
`tableReuseEnabled`, `renderTableFragment`, and the legacy table-paint selection.
Leave one `layoutTable` and one `paintTableLayout` production route.

- [ ] **Step 4: Verify Green and prove one route**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table-pagination.test.ts packages/docx/src/float-table-{geometry,page-fit,width}.test.ts packages/docx/src/pagination.test.ts
rg -n 'tableRequiresLegacyPaint|isFragmentPaintableTable|tableReuseEnabled|renderTableFragment' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: all tests pass and `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A5**

Commit subject: `refactor(docx): unify floating and split table layout`.
Use the roadmap review gate.

### Task A6: Extract the page/column state machine and establish worker parity

**Files:**

- Create: `packages/docx/src/layout/context.ts`
- Create: `packages/docx/src/layout/paginator.ts`
- Create: `packages/docx/src/layout/paginator.test.ts`
- Create: `packages/docx/src/layout/worker-parity.test.ts`
- Create: `packages/docx/src/document-layout-options.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/types.ts`
- Modify: `packages/docx/src/render-worker.ts`
- Modify: `packages/docx/src/worker-protocol.ts`
- Modify: `packages/docx/src/document.ts`
- Modify: `packages/docx/src/bookmark-nav.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: paragraph/table layout functions, A1 fingerprints, and A2 keyed options/convergence.
- Produces: `PageFlowState`, `PageFlowEvent`, final `layoutDocument(document, services, options)`, and worker-retained keyed `DocumentLayout` variants.

**Specification evidence:** ECMA-376 §17.6.4 (`w:cols`), §17.6.11
(`w:pgMar`), §17.6.13 (`w:pgSz`), §17.6.17 (`w:sectPr`), §17.6.22
(`w:type`), §17.6.12 (`w:pgNumType`), §17.3.1.15 (`w:keepNext`),
§17.3.1.14 (`w:keepLines`), §17.3.1.44 (`w:widowControl`), and explicit
run/paragraph break elements define page and column transitions.

```ts
export interface PageFlowState { readonly pageIndex: number; readonly columnIndex: number; readonly cursorYPt: number; readonly section: SectionLayoutContext }
export type PageFlowEvent =
  | Readonly<{ type: 'place'; node: PaintNode }>
  | Readonly<{ type: 'next-column' }>
  | Readonly<{ type: 'next-page'; reason: 'overflow' | 'explicit-break' | 'section-break' | 'parity' }>
  | Readonly<{ type: 'begin-section'; section: SectionLayoutContext }>;
export function paginateBody(input: BodyLayoutInput, services: LayoutServices, options: LayoutOptions): DocumentLayout;
```

- [ ] **Step 1: Add failing state-machine and parity tests**

Cover explicit page/column breaks, continuous and next-page sections, even/odd
parity pages, mixed page sizes, per-section vertical direction, multi-column
regions starting mid-page, keep-next, widow/orphan control, hidden paragraphs,
and bottom-margin overflow. Include DATE/TIME and NUMPAGES cases whose text changes
wrapping. Serialize the same synthetic document plus identical
`LayoutOptions`/font inventory through direct layout and the render-worker layout
handler and assert identical fingerprints and page sizes. Assert different
`currentDate` keys build isolated layout variants while the load-time default
metadata remains immutable and page validation uses the selected variant.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/paginator.test.ts packages/docx/src/layout/worker-parity.test.ts packages/docx/src/document-layout-options.test.ts packages/docx/src/pagination.test.ts packages/docx/src/per-section-headers-footers.test.ts
```

Expected: the worker retains `PaginatedBodyElement[][]`, and section/page facts
are recovered from element stamps rather than `LayoutPage`.

- [ ] **Step 3: Implement explicit transitions and page ownership**

Make each transition return a new `PageFlowState`; store section, geometry,
columns, content origin, page numbering, direction, header/footer references,
and parity-page metadata on `LayoutPage`. Replace `computePages` closure state
with `paginateBody`. The render worker retains `Map<string, DocumentLayout>` keyed
only by `layoutOptionsKey(options, services)`, where `services` is the actual
worker-owned instance; page metadata and bookmarks derive from the load-time default
layout, while render/collect requests carry the normalized key inputs needed to
select a variant. Keep worker protocol response shapes and public methods
unchanged; add only internal request fields.

Remove `sectionBreakSpacer`, `collapsedSpacer`, `leadsCollapsedRun`,
`hiddenCollapsed`, `colIndex`, `colGeom`, `colTopPt`, `sectionHF`,
`sectionGeom`, `sectionPageNumType`, and `sectionTextDirection` from
`PaginatedBodyElement`, then remove `PaginatedBodyElement` if no consumers remain.

- [ ] **Step 4: Verify Green and deletion**

Run:

```bash
pnpm vitest run packages/docx/src/layout/paginator.test.ts packages/docx/src/layout/worker-parity.test.ts packages/docx/src/document-layout-options.test.ts packages/docx/src/pagination.test.ts packages/docx/src/per-section-headers-footers.test.ts packages/docx/src/document-destroy.test.ts
rg -n 'sectionBreakSpacer|collapsedSpacer|leadsCollapsedRun|hiddenCollapsed|colGeom|colTopPt|sectionHF|sectionGeom|sectionPageNumType|sectionTextDirection|PaginatedBodyElement' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: tests pass; no runtime stamp or `PaginatedBodyElement` production match
remains; normalized main and worker fingerprints are equal.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A6**

Commit subject: `refactor(docx): make page flow an immutable state machine`.
Use the roadmap review gate.
