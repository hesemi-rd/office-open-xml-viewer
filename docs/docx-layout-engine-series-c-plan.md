# DOCX Layout Series C Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Finish compatibility, diagnostics, conformance, and cleanup infrastructure so the DOCX renderer has one auditable production algorithm.

**Architecture:** Floats use explicit constraints and the stable Series A convergence/service foundation; unsupported facts become diagnostics; a generated public corpus proves geometry and invariants. The final PR makes `renderer.ts` a thin compatibility adapter and statically proves no legacy path remains.

**Tech Stack:** TypeScript, Rust, Canvas 2D, FontFace/worker FontFaceSet, Vitest, Playwright, ast-grep, pnpm, GitHub Actions.

## Global Constraints

- Follow all constraints and the per-PR independent review gate in `docx-layout-engine-implementation-roadmap.md`.
- Compatibility rules require a normative citation, Microsoft implementation note, or documented synthetic Office observation.
- Non-convergence is an error diagnostic; it never returns stale or overlapping geometry.
- Generated fixtures must be redistributable and contain no private content.

---

### Task C1: Express float placement as explicit constraints with isolated compatibility

**Files:**

- Create: `packages/docx/src/layout/floats.ts`
- Create: `packages/docx/src/layout/floats.test.ts`
- Modify: `packages/docx/src/layout/compatibility.ts`
- Modify: `packages/docx/src/layout/compatibility.test.ts`
- Modify: `packages/docx/src/layout/paginator.ts`
- Modify: `packages/docx/src/layout/diagnostics.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/float-line-start-one-inch.test.ts`
- Modify: `packages/docx/src/float-table-geometry.test.ts`

**Interfaces:**

- Consumes: placed nodes, page/column bounds, wrap geometry, A1 diagnostics/fingerprints, and A2 `convergeLayout`.
- Produces: `FloatConstraint`, `FloatPlacement`, and `solveFloatPlacement`; it does not introduce another convergence engine.

**Specification evidence:** ECMA-376 §20.4.2.3 (`wp:anchor`),
§20.4.2.10/§20.4.2.11 positioning, §20.4.2.15–§20.4.2.20 wrap
geometry, `allowOverlap`, and `layoutInCell` define constraints.
[MS-OE376] §2.1.474 defines the Office `shapeLayoutLikeWW8` negative
line-relative-offset compatibility behavior. Each rule in
`compatibility.ts` names its Microsoft note or documented synthetic Office
observation; generic solver code contains no observation-derived constant.

```ts
export interface FloatConstraint { readonly anchor: SourceRef; readonly horizontal: AxisConstraint; readonly vertical: AxisConstraint; readonly wrap: WrapConstraint; readonly allowOverlap: boolean; readonly layoutInCell: boolean }
export interface FloatPlacement { readonly inkBounds: LayoutRect; readonly exclusion: readonly PointPt[]; readonly pageIndex: number; readonly columnIndex: number }
export function solveFloatPlacement(input: FloatSolveInput): FloatPlacement;
```

- [ ] **Step 1: Add failing constraint and convergence tests**

Cover page/margin/column/character anchors, align versus offset precedence,
square/tight/through/top-bottom wrap, `allowOverlap`, `layoutInCell`, negative
offsets and multiple interacting floats. Assert stable placement fingerprints,
then wrap the solver with A2 `convergeLayout` to assert repeated-fingerprint
cycle detection and `NON_CONVERGENCE` when `limit` is reached.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/floats.test.ts packages/docx/src/layout/compatibility.test.ts packages/docx/src/float-line-start-one-inch.test.ts packages/docx/src/float-table-geometry.test.ts
```

Expected: float placement mutates renderer state and lacks explicit cycle/error
contracts.

- [ ] **Step 3: Implement the pure solver and isolate compatibility**

Translate OOXML anchor facts to axis/wrap constraints, solve against immutable
page/container exclusions, and return placement plus exclusion polygon. Feed
interacting-float iterations through A2 `convergeLayout`; do not duplicate its
seen-set or safety-limit implementation.

Move only evidenced Office-specific behavior into `compatibility.ts`, with a
named function, evidence comment, and synthetic test per rule. Delete mutable
float placement/retry logic from `renderer.ts`.

- [ ] **Step 4: Verify Green and deterministic failure**

Run the command from Step 2 plus:

```bash
pnpm vitest run packages/docx/src/layout/paginator.test.ts packages/docx/src/layout/invariants.test.ts
pnpm typecheck
```

Expected: tests pass; identical input produces identical float fingerprints; an
oscillating fixture fails with `NON_CONVERGENCE` rather than stale geometry.

- [ ] **Step 5: Commit, independently review, fix, and merge PR C1**

Commit subject: `refactor(docx): solve floating layout as constraints`.
Use the roadmap review gate.

### Task C2: Propagate diagnostics and add a synthetic conformance corpus

**Files:**

- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Create: `packages/docx/parser/tests/diagnostics.rs`
- Modify: `packages/docx/src/parser-model.ts`
- Modify: `packages/docx/src/layout/diagnostics.ts`
- Create: `packages/docx/src/layout/diagnostics.test.ts`
- Create: `packages/docx/src/conformance/generate.ts`
- Create: `packages/docx/src/conformance/cases.ts`
- Create: `packages/docx/src/conformance/layout.test.ts`
- Create: `packages/docx/tests/visual/conformance.spec.ts`
- Modify: `packages/docx/playwright.config.ts`
- Modify: `.github/workflows/ci.yml`

**Interfaces:**

- Consumes: parser-preserved unsupported/invalid facts and final layout fingerprints.
- Produces: serialized `ParseDiagnostic`, mapped `LayoutDiagnostic`, generated minimal DOCX fixtures, and browser geometry assertions.

**Specification evidence:** Every generated case records its ECMA-376 element
and section or Microsoft implementation note beside the expected invariant.
Parser diagnostics distinguish schema-recognized unsupported content, invalid
values, and compatibility observations. A diagnostic never includes document
text or private source content and never changes the exported TypeScript model.

```rust
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParseDiagnostic {
    pub code: String,
    pub severity: DiagnosticSeverity,
    pub part: String,
    pub path: Vec<usize>,
}
```

Rust adds `diagnostics: Vec<ParseDiagnostic>` to its serialized document model.
The non-exported TypeScript parser boundary reads it through
`ParsedDocxDocumentModel = DocxDocumentModel & { diagnostics?:
ParseDiagnosticWire[] }`; `packages/docx/src/types.ts` and the A1 public
declaration baseline remain unchanged.

- [ ] **Step 1: Add failing parser-to-layout diagnostic tests**

Build minimal OOXML for a recognized unsupported decoration, invalid geometry,
unknown enum value, and a supported control case. Assert stable codes and source
paths; assert layout maps recoverable cases to warnings, fatal geometry to an
error, and the supported case to no diagnostic.

- [ ] **Step 2: Add failing generated-corpus geometry tests**

Generate redistributable pairwise cases spanning story, container, paragraph,
table, nested table, inline/floating object, direction, spacing, style source,
font source, and anchor reference. Assert page count, line ranges, non-overlap,
bottom-margin clearance, exact deterministic service fingerprints, and
main/worker parity in Vitest. In Chromium, Firefox, and WebKit assert the same
semantic invariants and exact main/worker parity within that browser; compare
cross-browser coordinates with explicit per-primitive tolerances because native
Canvas shaping may differ. Exact cross-browser fingerprints are required only if
the project later supplies one deterministic shaping engine and bundled font
corpus.

Use these explicit comparison rules: authored page boxes, fixed table grids,
fixed border endpoints, and drawing extents compare within `1e-6 pt`; text
`inkBounds` edges compare within `0.75 pt`; line/page partition equality is not a
cross-browser assertion when the browser supplies native shaping. Every browser
must still pass finite geometry, flow ownership, bottom-margin, structured-clone,
and exact same-browser main/worker fingerprint invariants.

- [ ] **Step 3: Run tests to verify Red**

Run:

```bash
cargo test -p docx-parser diagnostics -- --nocapture
pnpm vitest run packages/docx/src/layout/diagnostics.test.ts packages/docx/src/conformance/layout.test.ts
pnpm playwright test --config packages/docx/playwright.config.ts conformance.spec.ts --project=chrome
```

Expected: diagnostic fields and generated fixture modules do not exist.

- [ ] **Step 4: Preserve diagnostics and generate fixtures deterministically**

Record stable parser codes without document text, map them at the parser/layout
boundary, and include them in `DocumentLayout.diagnostics`. Implement a deterministic
ZIP/XML generator using repository dependencies, with fixed timestamps and IDs,
so generated bytes and deterministic-service expected fingerprints are stable.
Broaden the DOCX Playwright config to include the committed visual/conformance
test directory and add explicit Chrome, Firefox, and WebKit projects. CI runs
node geometry on every change and all three browser projects on the existing
browser-test cadence.

- [ ] **Step 5: Rebuild and verify Green**

Run:

```bash
pnpm build:wasm
cargo test -p docx-parser
pnpm vitest run packages/docx/src/layout/diagnostics.test.ts packages/docx/src/conformance/layout.test.ts
pnpm playwright test --config packages/docx/playwright.config.ts conformance.spec.ts
pnpm typecheck
```

Expected: all checks pass and two consecutive corpus generations have identical
hashes and deterministic-service fingerprints; browser tolerance and parity
assertions pass in all configured projects.

- [ ] **Step 6: Commit, independently review, fix, and merge PR C2**

Commit subject: `test(docx): add layout diagnostics and conformance corpus`.
Use the roadmap review gate.

### Task C3: Reduce renderer to an adapter and prove architectural completion

**Files:**

- Create: `packages/docx/src/paint/canvas-page.test.ts`
- Modify: `packages/docx/src/paint/canvas-page.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/document.ts`
- Delete: `packages/docx/src/layout-context.ts`
- Delete: `packages/docx/src/paragraph-measure.ts`
- Delete: `packages/docx/src/layout-fragments.ts`
- Delete: `packages/docx/src/table-fragments.ts`
- Create: `rules/no-docx-runtime-layout-stamps.yml`
- Create: `rules/no-docx-migration-flags.yml`
- Create: `rule-tests/no-docx-runtime-layout-stamps-test.yml`
- Create: `rule-tests/no-docx-migration-flags-test.yml`
- Create: `packages/docx/src/layout/architecture.test.ts`
- Modify: `scripts/check-docx-public-api.mjs`
- Modify: `scripts/check-docx-layout-boundaries.mjs`
- Modify: `sgconfig.yml`
- Modify: `.github/workflows/ci.yml`
- Modify: `docs/docx-layout-engine-redesign.md`

**Interfaces:**

- Consumes: final `layoutDocument`, `paintLayoutPage`, and layout-derived metadata.
- Produces: existing `paginateDocument` and `renderDocumentToCanvas` compatibility entry points as thin adapters with unchanged public callers.

**Specification evidence:** This PR introduces no OOXML behavior. It proves that
all normative and compatibility decisions are owned by the final layout modules,
that paint dependencies are measurement-free transitively, and that the public
declaration surface is byte-equivalent to the A1 baseline.

```ts
export function paginateDocument(document: Readonly<DocxDocumentModel>): DocumentLayout;
export async function renderDocumentToCanvas(document: Readonly<DocxDocumentModel>, target: HTMLCanvasElement | OffscreenCanvas, pageIndex: number, options: RenderPageOptions): Promise<void>;
```

- [ ] **Step 1: Write failing architecture tests and static rules**

Assert deep-frozen parser and layout inputs survive layout, repeated paint, failed
image paint, and search projection unchanged. Assert two layout calls with the
same services/options have identical fingerprints. Add tested static rules
rejecting runtime stamp properties, `*ReuseEnabled`, `*PaintEnabled`,
`RequiresLegacy`, legacy layout branches, dry-run layout, and layout/measurement
declarations. Extend the A1 import-graph checker so every paint entry's transitive
dependencies are free of measurement, shaping, style merge, pagination, and
parser-object access, and so every pagination/layout entry is on an explicit
allowlist outside `renderer.ts`.

Build `packages/docx/dist/types/index.d.ts`, normalize source-map paths/comments,
and compare it with `packages/docx/api/public-api-baseline.d.ts`. The comparison
must fail on any added, removed, or changed exported declaration; it replaces a
manual four-file diff.

- [ ] **Step 2: Run tests and static scan to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/architecture.test.ts packages/docx/src/paint/canvas-page.test.ts
pnpm lint
pnpm lint:test
node scripts/check-docx-layout-boundaries.mjs
pnpm --filter @silurus/ooxml-docx build
node scripts/check-docx-public-api.mjs
```

Expected: current renderer still contains mutable render/layout state and legacy
layout declarations, so the new assertions/rules fail.

- [ ] **Step 3: Complete adapter extraction and delete transitional modules**

Keep only resource acquisition, public-option normalization, service creation,
layout invocation/cache ownership, canvas sizing, and paint invocation in
`renderer.ts`. Move no algorithm into another compatibility wrapper. Delete each
transitional module only after `rg` proves no import remains. Update the design
status to implemented and record the final module dependency direction.

- [ ] **Step 4: Run focused architectural proof**

Run:

```bash
pnpm vitest run packages/docx/src/layout/architecture.test.ts packages/docx/src/paint/canvas-page.test.ts packages/docx/src/layout/invariants.test.ts packages/docx/src/layout/worker-parity.test.ts
pnpm lint
pnpm lint:test
node scripts/check-docx-layout-boundaries.mjs
rg -n 'fitMeasureReuseEnabled|fragmentPaintEnabled|lineReuseEnabled|tableReuseEnabled|RequiresLegacy|requiresLegacy|dryRun|PaginatedBodyElement|tableColWidthsPt|tableRowHeightsPt|layoutLinesInputs|deferFront' packages/docx/src --glob '!**/*.test.ts'
```

Expected: tests and static scan pass and `rg` has no production matches.

- [ ] **Step 5: Run the final broad verification and API diff**

Run:

```bash
pnpm build:wasm
pnpm test
pnpm typecheck
pnpm build-storybook
pnpm playwright test --config packages/docx/playwright.config.ts conformance.spec.ts
cargo test -p docx-parser
pnpm --filter @silurus/ooxml-docx build
node scripts/check-docx-public-api.mjs
node scripts/check-docx-layout-boundaries.mjs
git diff --check
```

Expected: all verification passes; generated declarations exactly match the A1
baseline; every paint dependency is layout/measurement-free; all layout and
pagination declarations are allowlisted in focused modules; no private artifact
is tracked.

- [ ] **Step 6: Obtain an independent final architecture audit**

Use the roadmap review brief and additionally require the reviewer to prove each
release-gate claim with a command or exact code reference. Expected: one production
algorithm per feature class, no paint measurement, no parser stamps, no migration
flags/fallback, all stories in layout, main/worker parity, and compatible public APIs.

- [ ] **Step 7: Fix findings, reverify, commit, and merge PR C3**

Commit subject: `refactor(docx): complete immutable layout pipeline`.
Repeat Steps 4–6 after material fixes, then merge with `gh pr merge <number> --merge`.
Close Issue #1037 only after GitHub shows PR C3 merged. Do not create a release or tag.
