# DOCX Layout Engine Redesign Implementation Roadmap

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the dual DOCX pagination/paint architecture with one immutable point-space layout result while preserving every public API.

**Architecture:** Rust preserves resolved OOXML facts; focused TypeScript layout modules produce a structured-clone-safe `DocumentLayout`; focused Canvas modules paint only stored geometry. The migration is divided into independently testable pull requests, and each migrated feature deletes its old measurement, stamp, gate, and fallback in the same pull request.

**Tech Stack:** Rust/wasm-bindgen, TypeScript, Canvas 2D/OffscreenCanvas, Vitest, Playwright, ast-grep, pnpm, GitHub Actions.

## Global Constraints

- Track the work in [Issue #1037](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1037).
- Keep `DocxViewer`, `DocxScrollViewer`, `DocxDocument`, render option signatures, and worker-mode public behavior backward compatible.
- Use points at scale 1 for all layout geometry; only paint reads scale or DPR.
- Follow ECMA-376 / ISO/IEC 29500 first, Microsoft implementation notes second, and isolated evidenced compatibility behavior third.
- Do not add sample-name branches, fixture constants, empirical fit factors, production migration flags, or a silent legacy fallback.
- Do not commit private samples, private stories, local VRT references, generated WASM, or local absolute paths in public text.
- Rebuild WASM before parser-backed integration tests whenever Rust parser sources change.
- Do not release any package or tag until Series A, B, and C and the final audit are merged.
- Every PR uses TDD, removes the migrated legacy path, receives an independent critical agent review, fixes valid findings, reruns verification, and merges with a merge commit.

---

## File and ownership map

The final source tree has one responsibility per module:

```text
packages/docx/src/layout/
  types.ts         immutable layout result and source-reference types
  diagnostics.ts   diagnostic codes and collector
  context.ts       document/section/story/container resolution
  flow.ts          block dispatch and recursive container coordination
  text.ts          font selection, shaping, and measured glyph contracts
  resources.ts     image/math metadata and stable resource keys
  convergence.ts   fingerprint/cycle/limit convergence primitive
  compatibility.ts evidenced Office-specific behavior only
  error-page.ts    parse-error page layout through the text service
  paragraph.ts     paragraph line layout and decoration geometry
  table.ts         columns, rows, cells, continuations, and border segments
  floats.ts        wrap constraints, placement, and convergence
  stories.ts       body/header/footer/note/text-box block layout
  paginator.ts     explicit page/column state machine
  invariants.ts    deterministic validation and normalized fingerprints

packages/docx/src/paint/
  canvas-page.ts    page layer ordering and transforms
  canvas-text.ts    glyph/decorations paint from TextPlacement[]
  canvas-table.ts   resolved cell backgrounds and border segments
  canvas-drawing.ts drawing/image/shape paint from resolved bounds
```

Existing shared DrawingML, color, image, and font primitives remain in
`packages/core`; DOCX page flow stays in `packages/docx`. `renderer.ts` becomes a
resource-preparation adapter and exports compatibility entry points only.

A1 records an explicit inventory of `packages/core` and
`packages/ooxml-common` before introducing geometry, transform, clip, image,
font, diagnostic, or Canvas contracts. Shared format-neutral primitives are
reused there; WordprocessingML flow contracts remain DOCX-local.

## Stable internal interfaces

Series work must converge on these exact contracts; later plans consume these
names rather than inventing parallel representations:

```ts
export type LayoutNodeId = string;
export type SourceRef = Readonly<{
  story: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textbox';
  storyInstance: string;
  path: readonly number[];
}>;

export type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
  : T extends readonly (infer U)[] ? readonly DeepReadonly<U>[]
  : T extends object ? { readonly [K in keyof T]: DeepReadonly<T[K]> }
  : T;

export interface Matrix2DData {
  readonly a: number; readonly b: number; readonly c: number;
  readonly d: number; readonly e: number; readonly f: number;
}
export type ClipPathData =
  | Readonly<{ kind: 'rect'; rect: LayoutRect }>
  | Readonly<{ kind: 'polygon'; points: readonly PointPt[] }>;

export interface LayoutServices {
  readonly text: TextLayoutService;
  readonly images: ImageMetadataService;
  readonly math: MathMetadataService;
}

export interface LayoutOptions {
  readonly currentDateMs: number;
}

export function layoutOptionsKey(options: LayoutOptions, services: LayoutServices): string;

export interface DocumentLayout {
  readonly pages: readonly LayoutPage[];
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export interface LayoutPage {
  readonly pageIndex: number;
  readonly geometry: PageGeometry;
  readonly section: SectionLayoutContext;
  readonly layers: PageLayers;
  readonly readingOrder: readonly LayoutNodeId[];
}

export interface PageLayers {
  readonly paintOrder: readonly PagePaintEntry[];
  readonly background: readonly PaintNode[];
  readonly behindText: readonly PaintNode[];
  readonly header: readonly PaintNode[];
  readonly body: readonly PaintNode[];
  readonly notes: readonly PaintNode[];
  readonly front: readonly PaintNode[];
  readonly footer: readonly PaintNode[];
}

export type PageLayerId = 'background' | 'behindText' | 'header' | 'body' | 'notes' | 'front' | 'footer';
export interface PagePaintEntry { readonly layer: PageLayerId; readonly nodeId: LayoutNodeId }

export function layoutDocument(
  document: Readonly<DocxDocumentModel>,
  services: LayoutServices,
  options: LayoutOptions,
): DocumentLayout;

export async function paintLayoutPage(
  layout: DocumentLayout,
  pageIndex: number,
  target: HTMLCanvasElement | OffscreenCanvas,
  options: PaintPageOptions,
): Promise<void>;
```

`PaintNode` is the union of `ParagraphLayout`, `TableLayout`, `DrawingLayout`,
`TextBoxLayout`, and `NoteLayout`. Every flow node distinguishes `flowBounds`
(pagination ownership), `inkBounds` (glyph/border overhang), optional
`clipBounds`, and `advancePt`. Overlap and bottom-margin invariants apply to
ordinary in-flow `flowBounds`, never to allowed floating overlap, negative
spacing ink, frames, or clipped overhang. Each node contains all text, resolved
font, color, border, transform, clipping, and resource-key data needed by paint;
paint never dereferences parser objects.

`LayoutOptions.currentDateMs` is normalized once. Each text/image/math service
owns a readonly fingerprint derived from its immutable resource snapshot; callers
cannot supply those strings independently. Main and worker retain layouts by
`layoutOptionsKey(options, services)`; `NUMPAGES` is solved inside convergence, not supplied as a
paint option. The load-time default option key determines the synchronous
`pageCount`/`pageSize` getters. A per-call `currentDate` selects or lazily builds
a keyed layout variant for that render/collection request and validates its
requested page against that variant; it does not mutate the default metadata or
public API. Contract tests document this existing getter limitation explicitly.

## Pull request sequence

| PR | Deliverable | Deletes in the same PR | Depends on |
|---|---|---|---|
| Plan | Approved design and executable plans | None | Issue #1037 |
| A1 | Layout result, deep-readonly/plain-data types, shared-primitives audit, invariants, and enforceable boundary gates | None; the transitional fragment rule remains until A3 | Plan |
| A2 | Stable font/text/image/math services, layout options/cache keys, convergence, and parse-error layout | global font state and parse-error paint measurement | A1 |
| A3 | All body paragraphs and every `DocRun` variant use self-contained layout nodes | paragraph line stamps, reuse flags, paragraph fragment gates, transitional fragment rule/test | A2 |
| A4 | One measurement builds in-flow and nested `TableLayout` | table width/height stamps and second cell measurement for this class | A3 |
| A5 | Floating and split tables use the same table result | `tableRequiresLegacyPaint`, legacy table gate, remaining table flags | A4 |
| A6 | Explicit page/column state machine, keyed layout cache, and main/worker fingerprint parity | page/section stamps on `PaginatedBodyElement`, page-flow closures | A5 |
| B1 | Headers, footers, footnotes, and endnotes use shared story layout and A2 convergence | story dry-run measurement and note paint relayout | A6 |
| B2 | Parser preserves complete `txbxContent` on an internal wire model; text boxes use shared block layout | text-box-specific paragraph engine; exported `textBlocks` remains as a compatibility projection | B1 |
| B3 | `PageLayers` owns drawing order | `deferFront` callback side channel | B2 |
| B4 | Search/selection geometry projects from layout | `onTextRun` paint callbacks and `collectRuns` dry render | B3 |
| C1 | Float solver uses explicit constraints and the A2 convergence primitive; compatibility rules are isolated | hidden float retry/placement mutation | B4 |
| C2 | Parser/layout diagnostics and synthetic conformance corpus | silent unsupported-feature omission | C1 |
| C3 | Thin renderer adapter, public declaration compatibility, and transitive architecture audit | all remaining layout code, stamps, flags, and fallback gates in `renderer.ts` | C2 |

The executable Red/Green/deletion steps are in:

- `docs/docx-layout-engine-series-a-plan.md`
- `docs/docx-layout-engine-series-b-plan.md`
- `docs/docx-layout-engine-series-c-plan.md`

## Per-PR review and merge gate

After the implementation and local verification of every PR:

- [ ] **Step 1: Commit only intended public source, tests, rules, and docs**

Use a repository-style English commit subject and a body containing the root
cause, normative section or observed compatibility boundary, and verification.
Append the repository-required co-author trailer using the actual implementing
model. If orchestration or implementation used multiple model providers, include
the actual matching trailer for each contributor; never copy a hard-coded model
name from this roadmap.

- [ ] **Step 2: Run broad verification**

Run:

```bash
pnpm test
pnpm typecheck
pnpm lint
pnpm lint:test
pnpm test:docx-boundaries
node scripts/check-docx-layout-boundaries.mjs
pnpm --filter @silurus/ooxml-docx build
node scripts/check-docx-public-api.mjs
pnpm build-storybook
git diff --check origin/main...HEAD
```

Expected: all commands exit 0. Parser PRs additionally run
`cargo test -p docx-parser` after `pnpm build:wasm`.

- [ ] **Step 3: Obtain an independent critical review**

Give a fresh agent the base commit, head commit, design document, current series
plan, and this exact review brief:

```text
Review this PR critically. Verify OOXML spec-first behavior and cited evidence,
single responsibility, duplicated layout/measurement/style/paint logic,
appropriate sharing versus false abstraction across docx/xlsx/pptx/core,
public API compatibility, worker parity, parser/layout immutability, deterministic
geometry, absence of fixture heuristics and silent fallback, and behavior-level
test quality. Inspect the whole base..head diff. Report actionable findings with
severity and exact files/lines; do not assume the implementation is correct.
```

Expected: a written review covering every category, including an explicit “no
findings” statement for categories without findings.

- [ ] **Step 4: Resolve every valid finding and repeat verification**

Apply fixes with TDD, commit the review-fix source and tests, rerun the focused
failing test plus all commands from Step 2, and ask the independent agent to
re-review when a fix changes an interface or algorithm. Expected: no unresolved
valid findings and no uncommitted PR changes.

- [ ] **Step 5: Publish and merge**

Push the `codex/` branch, open a PR linked to `#1037`, wait for required checks,
and run `gh pr merge <number> --merge`. Expected: GitHub reports the PR merged by
a merge commit; never squash and never push directly to `main`.

## Plan PR task

### Task 1: Land the approved design and executable roadmap

**Files:**

- Modify: `docs/docx-layout-engine-redesign.md`
- Create: `docs/docx-layout-engine-implementation-roadmap.md`
- Create: `docs/docx-layout-engine-series-a-plan.md`
- Create: `docs/docx-layout-engine-series-b-plan.md`
- Create: `docs/docx-layout-engine-series-c-plan.md`

**Interfaces:**

- Consumes: the approved architecture in `docs/docx-layout-engine-redesign.md` and Issue #1037.
- Produces: the stable interface names and PR dependency order above.

- [ ] **Step 1: Verify every design requirement maps to a PR**

Run:

```bash
rg -n '^## |^### ' docs/docx-layout-engine-redesign.md docs/docx-layout-engine-*-plan.md docs/docx-layout-engine-implementation-roadmap.md
```

Expected: parser facts, contexts, layout, paint, convergence, worker mode,
invariants, synthetic corpus, review gate, and release gate all appear in the
plans.

- [ ] **Step 2: Reject placeholders and unsafe fixture language**

Run:

```bash
rg -n 'T[B]D|TO[D]O|implement [l]ater|similar [t]o|s[a]mple-[0-9]+|p[r]ivate/' docs/docx-layout-engine-*-plan.md docs/docx-layout-engine-implementation-roadmap.md
```

Expected: no matches.

- [ ] **Step 3: Verify plan paths and current symbols**

Run:

```bash
for p in packages/docx/src packages/docx/parser/src rules .github/workflows; do test -e "$p"; done
rg -n 'fitMeasureReuseEnabled|fragmentPaintEnabled|lineReuseEnabled|tableReuseEnabled|tableRequiresLegacyPaint|deferFront|collectRuns' packages/docx/src
```

Expected: all roots exist and each named legacy symbol currently has at least one
match, proving the deletion checks are grounded in the current source.

- [ ] **Step 4: Check formatting**

Run: `git diff --check`

Expected: exit 0.

- [ ] **Step 5: Commit**

```bash
git add docs/docx-layout-engine-redesign.md docs/docx-layout-engine-implementation-roadmap.md docs/docx-layout-engine-*-plan.md
git commit -m "docs(docx): plan immutable layout engine migration"
```

Expected: one documentation commit containing no private fixture content.
