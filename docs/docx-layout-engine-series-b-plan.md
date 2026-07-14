# DOCX Layout Series B Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Put headers, footers, notes, text boxes, drawings, and search geometry into the same immutable page layout used by the body.

**Architecture:** Story kind and container kind remain orthogonal: every story lays out block content through the same paragraph/table functions, while page layers own physical placement and paint order. Parser changes preserve authored block structure instead of flattening content for a renderer-specific shortcut.

**Tech Stack:** Rust/quick-xml/wasm-bindgen, TypeScript, Canvas 2D, OffscreenCanvas, Vitest, pnpm.

## Global Constraints

- Follow all constraints and the per-PR independent review gate in `docx-layout-engine-implementation-roadmap.md`.
- Rebuild DOCX WASM after parser changes.
- Preserve worker request/response and public method signatures.
- A story may contain paragraphs, tables, nested tables, and supported drawings through one shared block layout function.

---

### Task B1: Migrate headers, footers, footnotes, and endnotes to shared story layout

**Files:**

- Create: `packages/docx/src/layout/stories.ts`
- Create: `packages/docx/src/layout/stories.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/layout/paginator.ts`
- Modify: `packages/docx/src/paint/canvas-page.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/header-reserve.test.ts`
- Modify: `packages/docx/src/footer-reserve.test.ts`
- Modify: `packages/docx/src/footnote-in-table.test.ts`
- Modify: `packages/docx/src/per-section-headers-footers.test.ts`
- Modify: `packages/docx/src/vertical-header-footer.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: `layoutParagraph`, `layoutTable`, `LayoutPage`, and `PageLayers` from Series A.
- Produces: `StoryLayoutInput`, `StoryLayout`, `NoteLayout`, `layoutStory`, and converged note reservation.

**Specification evidence:** ECMA-376 §17.10 header/footer references and
selection (`w:headerReference`, `w:footerReference`, `w:titlePg`, even/odd
settings) and §17.11 footnote/endnote references, numbering, separators, and
placement define story ownership. This PR adds no observation-only reserve
heuristic; behavior not determined by those facts emits a diagnostic and requires
a separately evidenced compatibility rule.

```ts
export interface StoryLayout { readonly story: SourceRef['story']; readonly flowBounds: LayoutRect; readonly inkBounds: LayoutRect; readonly clipBounds?: LayoutRect; readonly blocks: readonly PaintNode[]; readonly advancePt: number }
export interface NoteLayout { readonly kind: 'note'; readonly id: LayoutNodeId; readonly source: SourceRef; readonly flowBounds: LayoutRect; readonly inkBounds: LayoutRect; readonly separator: readonly BorderSegment[]; readonly story: StoryLayout }
export function layoutStory(input: StoryLayoutInput, services: LayoutServices): StoryLayout;
```

- [ ] **Step 1: Add failing story and reservation tests**

Test first/default/even headers and footers, title pages, mixed sections, vertical
sections, tall header/footer reserve, footnote references in body and table cells,
multiple notes in reference order, and endnotes at document end. Assert layer,
bounds, reservation, and reading order, and ensure body/notes/footer rectangles do
not overlap.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/stories.test.ts packages/docx/src/header-reserve.test.ts packages/docx/src/footer-reserve.test.ts packages/docx/src/footnote-in-table.test.ts packages/docx/src/per-section-headers-footers.test.ts packages/docx/src/vertical-header-footer.test.ts
```

Expected: story nodes are absent and current code uses dry-run render states to
measure story height.

- [ ] **Step 3: Use one block story layout and explicit reservation convergence**

Resolve the active header/footer references from `LayoutPage.section`, lay out
their blocks through `layoutStory`, and place them into header/footer layers.
Collect note references from retained node metadata, lay out note stories, reduce
the body band, and call A2 `convergeLayout` until the relevant page fingerprint is
unchanged. Store note separator geometry on `NoteLayout`; do not create a second
note-specific convergence loop.

Delete `renderHeaderFooter`, `measureFootnoteHeight`, note dry-run render states,
and note paint-time block layout. No story paint path may call body layout.

- [ ] **Step 4: Verify Green and one story engine**

Run the command from Step 2, then:

```bash
rg -n 'renderHeaderFooter|measureFootnoteHeight|currentNoteNumber' packages/docx/src/renderer.ts
pnpm typecheck
```

Expected: tests pass and the old story measurement symbols have no production
matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR B1**

Commit subject: `refactor(docx): lay out page stories through one engine`.
Use the roadmap review gate.

### Task B2: Preserve complete text-box content and use shared block layout

**Files:**

- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Create: `packages/docx/src/parser-model.ts`
- Create: `packages/docx/src/parser-model.test.ts`
- Create: `packages/docx/src/layout/textbox.test.ts`
- Modify: `packages/docx/src/layout/textbox-input.ts`
- Modify: `packages/docx/src/layout/textbox-input.test.ts`
- Modify: `packages/docx/src/layout/stories.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/paint/canvas-drawing.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/textbox-main-engine.test.ts`
- Modify: `packages/docx/src/textbox-vertical-render.test.ts`
- Modify: `packages/docx/src/renderer.textbox-image.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: parser `BodyElement`, shared `layoutStory`, and DrawingML shape geometry.
- Produces: Rust wire-only `ShapeRun.text_box_content: Vec<TextBoxBlockWire>`, internal `ParsedShapeRun.textBoxContent?: Array<BodyElement | ParsedUnsupportedTextBoxBlock>`, the unchanged public `ShapeRun.textBlocks?: ShapeText[]`, and `TextBoxLayout`.

**Specification evidence:** ECMA-376 `wps:txbx/w:txbxContent` and
`wps:bodyPr`, ECMA-376 §20.4.2.38 (`w:txbxContent`), and
DrawingML §21.1.2.1.1 body insets/autofit define preserved block order and text
container geometry. Supported paragraph/table blocks are preserved in document
order; every other schema-permitted `EG_BlockLevelElts` kind is retained as an
internal unsupported-block marker with its QName/source path and emits a
diagnostic rather than being flattened or discarded silently.

```ts
export interface ParsedUnsupportedTextBoxBlock { readonly type: 'unsupportedTextBoxBlock'; readonly qName: string; readonly sourcePath: readonly number[] }
export interface TextBoxLayout { readonly kind: 'textbox'; readonly id: LayoutNodeId; readonly source: SourceRef; readonly flowBounds: LayoutRect; readonly inkBounds: LayoutRect; readonly clip: ClipPathData; readonly transform: Matrix2DData; readonly story: StoryLayout }
```

- [ ] **Step 1: Add failing parser and layout tests**

Build minimal OOXML containing `wps:wsp/wps:txbx/w:txbxContent` with paragraph,
table, nested table, image-bearing paragraph, and trailing paragraph. Assert the
Rust JSON preserves block order in wire-only `textBoxContent`, continues emitting
the current `textBlocks` compatibility projection, and the TypeScript layout
places all blocks within text-box insets for horizontal and vertical flow.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
cargo test -p docx-parser textbox -- --nocapture
pnpm vitest run packages/docx/src/layout/textbox.test.ts packages/docx/src/textbox-main-engine.test.ts packages/docx/src/textbox-vertical-render.test.ts packages/docx/src/renderer.textbox-image.test.ts
```

Expected: parser assertions fail because the current `ShapeTextBody` tuple and
`ShapeRun.text_blocks` preserve only the flattened `ShapeText` projection.

- [ ] **Step 3: Preserve blocks and delegate to `layoutStory`**

Replace the private Rust tuple alias with a named `ShapeTextBody` struct carrying
both `legacy_blocks: Vec<ShapeText>` and `content: Vec<TextBoxBlockWire>`. Define
the parse-time wire enum without using a Rust keyword as a field name:

```rust
#[derive(Serialize, Debug, Clone)]
#[serde(untagged)]
enum TextBoxBlockWire {
    Body(BodyElement),
    Unsupported {
        #[serde(rename = "type")]
        kind: String,
        #[serde(rename = "qName")]
        q_name: String,
        #[serde(rename = "sourcePath")]
        source_path: Vec<usize>,
    },
}
```

Set `kind` to the fixed value `"unsupportedTextBoxBlock"`. Build this enum while
walking `txbxContent`, so an unsupported block's QName, source path, and order are
retained before serialization rather than reconstructed from a narrowed
`BodyElement`. Add a
`#[serde(skip_serializing_if = "Vec::is_empty")] text_box_content:
Vec<TextBoxBlockWire>` wire field to Rust `ShapeRun` while retaining and populating
`text_blocks` exactly as today. Do not add, remove, or rename any exported
TypeScript declaration. `parser-model.ts` defines non-exported
`ParsedShapeRun = ShapeRun & { textBoxContent?: Array<BodyElement |
ParsedUnsupportedTextBoxBlock> }` and reads the
wire field; externally constructed models without it fall back to the existing
`textBlocks` compatibility projection.

Parse supported paragraph/table `txbxContent` blocks in document order through
existing body-element parsers and emit an unsupported-block marker for every
other schema-permitted block. Create a text-box container context containing insets,
autofit, vertical direction, clipping, and plain-data transform, then call
`layoutStory`. Paint the nested story from stored geometry. Extend A3's single
`normalizeTextBoxInput` boundary to prefer internal `ParsedShapeRun.textBoxContent`
and otherwise convert public `ShapeRun.textBlocks`. Generalize its internal return
type from paragraph-only input to the `StoryBlockInput` union used by
`layoutStory`; do not add a second adapter or delete this fallback after parser
wiring. Add a test that constructs only the exported public `ShapeRun` shape with
`textBlocks` (no parser-only fields) and asserts the same retained paragraph
layout as the A3 path. A3 already removed text-box-specific measurement and
`renderShapeText`. Retain the exported `ShapeText` and `ShapeTextRun` contracts.

- [ ] **Step 4: Rebuild and verify Green**

Run:

```bash
pnpm build:wasm
cargo test -p docx-parser
pnpm vitest run packages/docx/src/layout/textbox.test.ts packages/docx/src/textbox-*.test.ts packages/docx/src/anchored-textbox-render.test.ts packages/docx/src/renderer.textbox-image.test.ts
rg -n 'type ShapeTextBody = \(' packages/docx/src packages/docx/parser/src
pnpm typecheck
```

Expected: all checks pass, the public declaration baseline from A1 is unchanged,
`textBlocks` remains serialized, the external-public-model fallback test passes,
and `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR B2**

Commit subject: `refactor(docx): preserve text box block layout`.
Use the roadmap review gate.

### Task B3: Materialize drawing z-order in PageLayers

**Files:**

- Modify: `packages/docx/src/layout/types.ts`
- Create: `packages/docx/src/layout/page-layers.test.ts`
- Modify: `packages/docx/src/paint/canvas-page.ts`
- Modify: `packages/docx/src/paint/canvas-drawing.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/front-float-zorder.test.ts`
- Modify: `packages/docx/src/section-break-anchored-textbox-render.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: all self-contained paint nodes from A/B1/B2.
- Produces: `buildPageLayers` and final `paintLayoutPage` ordering.

**Specification evidence:** ECMA-376 §20.4.2.3 specifies that
`wp:anchor/@relativeHeight` orders objects with the same `behindDoc` value and
that every `behindDoc=false` object is above every `behindDoc=true` object. A
committed decision table distinguishes
body-, header-, footer-, note-, and text-box-owned anchors; where those sources
do not specify cross-story ordering, a reproducible synthetic Office observation
is marked as
compatibility evidence rather than presented as normative.

```ts
export function buildPageLayers(nodes: readonly PlacedLayoutNode[]): PageLayers;
export async function paintLayoutPage(layout: DocumentLayout, pageIndex: number, target: HTMLCanvasElement | OffscreenCanvas, options: PaintPageOptions): Promise<void>;
```

- [ ] **Step 1: Add failing layer-order tests**

Create interleaved behind-text and front drawings anchored in body, header,
footer, notes, and text boxes. Derive each expected `paintOrder` from the committed
normative/compatibility decision table, record Canvas operations, and assert
cross-story order, stable relative-height/document order within an equivalent
stacking context, and one paint per node. Do not encode one assumed global order
before the evidence table is complete.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/page-layers.test.ts packages/docx/src/front-float-zorder.test.ts packages/docx/src/section-break-anchored-textbox-render.test.ts
```

Expected: current `deferFront` callbacks determine order during paint and cannot
be inspected as layout data.

- [ ] **Step 3: Classify drawings during layout and paint fixed layers**

Use OOXML `behindDoc`, relative height/z-order, story ownership, and the decision
table to assign each placed drawing once. Materialize both semantic arrays and a
`PageLayers.paintOrder: readonly PagePaintEntry[]` sequence. Make
`canvas-page.ts` traverse that stored
sequence and dispatch by node kind. Delete `RenderState.deferFront`, callback
queues, and callback re-entry guards.

- [ ] **Step 4: Verify Green and callback deletion**

Run the command from Step 2, then:

```bash
rg -n 'deferFront|deferredFront' packages/docx/src
pnpm typecheck
```

Expected: tests pass and no production matches remain.

- [ ] **Step 5: Commit, independently review, fix, and merge PR B3**

Commit subject: `refactor(docx): make page layers own drawing order`.
Use the roadmap review gate.

### Task B4: Project search and selection geometry from DocumentLayout

**Files:**

- Create: `packages/docx/src/layout/text-index.ts`
- Create: `packages/docx/src/layout/text-index.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/document.ts`
- Modify: `packages/docx/src/render-worker.ts`
- Modify: `packages/docx/src/worker-protocol.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/find.test.ts`
- Modify: `packages/docx/src/find-highlight-layer.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`

**Interfaces:**

- Consumes: `TextPlacement` metadata retained by A3 and worker-retained keyed `DocumentLayout` from A6.
- Produces: `textRunsForPage` returning the existing public `DocxTextRunInfo[]` shape.

**Specification evidence:** ECMA-376 §17.13.6.2 bookmarks,
§17.16.22 hyperlinks, bidirectional/vertical run properties, and the
stored glyph-cluster mapping determine reading ranges and transformed hit
geometry. Selection rectangles are viewer projections, not OOXML layout rules;
tests therefore derive them from retained clusters without re-shaping.

```ts
export function textRunsForPage(layout: DocumentLayout, pageIndex: number, options: Readonly<{ scale: number }>): DocxTextRunInfo[];
```

- [ ] **Step 1: Write failing projection parity tests**

Cover ligatures, bidi visual order, tabs, hyperlinks, rotated/vertical text,
table cells, notes, headers, and text boxes. Assert projected runs match stored
text ranges and transformed bounds and that `collectPageRuns` performs zero
Canvas calls in main and worker modes.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/text-index.test.ts packages/docx/src/find.test.ts packages/docx/src/find-highlight-layer.test.ts
```

Expected: `collectPageRuns` creates an OffscreenCanvas and dry-renders the page.

- [ ] **Step 3: Implement layout projection without protocol breakage**

Traverse `readingOrder` and each node's `TextPlacement` records, apply stored page
transforms, and map to `DocxTextRunInfo`. Keep `collectPageRuns` and worker
`collectRuns` message names and response shapes unchanged, but satisfy them from
retained layout. Remove `RenderState.onTextRun`, paint callbacks, throwaway
canvases, and dry-render run collection.

- [ ] **Step 4: Verify Green and no dry render**

Run the command from Step 2, then:

```bash
rg -n 'onTextRun|new OffscreenCanvas\(1, 1\).*collect|collectRuns.*renderDocumentToCanvas' packages/docx/src
pnpm typecheck
```

Expected: tests pass; `onTextRun` remains only at the public adapter boundary,
not in paint/layout state; collect-runs does not render.

- [ ] **Step 5: Commit, independently review, fix, and merge PR B4**

Commit subject: `refactor(docx): project text geometry from layout`.
Use the roadmap review gate.
