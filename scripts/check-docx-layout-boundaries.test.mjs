import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const checker = resolve(import.meta.dirname, 'check-docx-layout-boundaries.mjs');

function write(root, path, contents) {
  const absolute = join(root, path);
  mkdirSync(resolve(absolute, '..'), { recursive: true });
  writeFileSync(absolute, contents);
}

function command(root, executable, args) {
  const result = spawnSync(executable, args, { cwd: root, encoding: 'utf8' });
  return {
    status: result.status,
    output: `${result.stdout ?? ''}${result.stderr ?? ''}`,
  };
}

function git(root, ...args) {
  const result = command(root, 'git', args);
  assert.equal(result.status, 0, result.output);
}

function runChecker(root, ...args) {
  return command(root, process.execPath, [checker, '--root', root, ...args]);
}

function initializeRepository() {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-'));
  write(root, 'packages/docx/src/renderer.ts', 'function buildMeasureState(ctx: unknown, fonts: unknown) { return [ctx, fonts]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}) { const measure = buildMeasureState(ctx, resolvedLocalFonts); return [measure]; }\nexport function computeTableLayout() { return []; }\n');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paint() {}\n');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base');
  git(root, 'switch', '-c', 'a1');
  return root;
}

function establishA1Baseline(root) {
  const writeResult = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');
  assert.equal(writeResult.status, 0, writeResult.output);
  const checkResult = runChecker(root, '--base-ref', 'main');
  assert.equal(checkResult.status, 0, checkResult.output);
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'establish boundary');
  git(root, 'switch', 'main');
  git(root, 'merge', '--ff-only', 'a1');
  git(root, 'switch', '-c', 'a2');
}

test('rejects a transitive paint edge to a measurement module', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { helper } from './helper.js';\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/helper.ts', "import { layoutLines } from '../line-layout.js';\nexport const helper = layoutLines;\n");
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*helper\.ts.*line-layout\.ts/s);
});

test('rejects any paint runtime dependency outside the paint owner directory', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-arbitrary-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { helper } from '../text-wrap.js';\nexport { helper };\n");
  write(root, 'packages/docx/src/text-wrap.ts', 'export const helper = true;\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*text-wrap\.ts/s);
});

test('allows only the layout contract as a type-only paint dependency', () => {
  const allowed = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-type-edge-'));
  write(allowed, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(allowed, 'packages/docx/src/layout/types.ts', 'export interface Layout { pages: number; }\n');
  write(allowed, 'packages/docx/src/paint/canvas-page.ts', "import type { Layout } from '../layout/types.js';\nexport type Page = Layout;\n");
  assert.equal(runChecker(allowed, '--final').status, 0);

  write(allowed, 'packages/docx/src/layout/flow.ts', 'export interface Flow { y: number; }\n');
  write(allowed, 'packages/docx/src/paint/canvas-page.ts', "import type { Flow } from '../layout/flow.js';\nexport type Page = Flow;\n");
  const forbidden = runChecker(allowed, '--final');
  assert.notEqual(forbidden.status, 0);
  assert.match(forbidden.output, /FORBIDDEN_PAINT_EDGE/);
});

test('allows only named shared atomic painters from core', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-shared-paint-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { renderChart } from '@silurus/ooxml-core';\nexport { renderChart };\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { measureTextWidth as renderChart } from '@silurus/ooxml-core';\nexport { renderChart };\n");
  const result = runChecker(root, '--final');
  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
});

test('audits dependencies of the retained page graph allowed in paint', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-page-graph-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "import { orderedPagePaintNodes } from '../layout/page-graph.js';\nexport { orderedPagePaintNodes };\n");
  write(root, 'packages/docx/src/layout/page-graph.ts', "import { measureTextWidth } from '../measurement.js';\nexport const orderedPagePaintNodes = measureTextWidth;\n");
  write(root, 'packages/docx/src/measurement.ts', 'export function measureTextWidth() { return 1; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*page-graph\.ts.*measurement\.ts/s);
});

test('rejects non-literal dynamic paint imports', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-dynamic-edge-'));
  write(root, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export const load = (name: string) => import(`../${name}.js`);\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /NON_LITERAL_MODULE_EDGE/);
});

test('rejects computed measurement access in paint and display inputs in layout TSX', () => {
  const paintRoot = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-computed-'));
  write(paintRoot, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(paintRoot, 'packages/docx/src/paint/canvas-page.ts', "export const width = (ctx: CanvasRenderingContext2D) => ctx['measureText']('x').width;\n");
  const paintResult = runChecker(paintRoot, '--final');
  assert.notEqual(paintResult.status, 0);
  assert.match(paintResult.output, /PAINT_CAPABILITY/);

  const layoutRoot = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-layout-display-'));
  write(layoutRoot, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(layoutRoot, 'packages/docx/src/layout/page.tsx', 'export const Page = ({ dpr }: { dpr: number }) => dpr;\n');
  const layoutResult = runChecker(layoutRoot, '--final');
  assert.notEqual(layoutResult.status, 0);
  assert.match(layoutResult.output, /LAYOUT_DISPLAY_CAPABILITY/);
});

test('rejects a CommonJS require edge that bypasses static ESM imports', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-require-'));
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "const helper = require('./helper.js');\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/helper.ts', "const measured = require('../line-layout.js');\nexport { measured };\n");
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FORBIDDEN_PAINT_EDGE/);
});

test('writes the A1 baseline only when the merge base has none', () => {
  const root = initializeRepository();

  const result = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
  assert.match(readFileSync(join(root, 'scripts/docx-layout-boundary-baseline.json'), 'utf8'), /computePages/);
  assert.equal(runChecker(root, '--base-ref', 'main').status, 0);
});

test('rejects a head baseline that expands the merge-base allowances', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  baseline.legacySymbolCounts.tableReuseEnabled = 1;
  writeFileSync(baselinePath, `${JSON.stringify(baseline, null, 2)}\n`);

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BASELINE_EXPANSION/);
  assert.match(result.output, /tableReuseEnabled/);
});

test('rejects moving a legacy symbol to another file without increasing its global count', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/legacy-copy.ts', 'export function computePages() { return []; }\n');

  const result = runChecker(root, '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /BASELINE_EXPANSION/);
});

test('rejects renaming a migration flag and changing a hash-frozen leaf declaration', () => {
  const renamed = initializeRepository();
  establishA1Baseline(renamed);
  write(renamed, 'packages/docx/src/new-switch.ts', 'export const useLegacyLayout = true;\n');
  const renamedResult = runChecker(renamed, '--base-ref', 'main');
  assert.notEqual(renamedResult.status, 0);
  assert.match(renamedResult.output, /BASELINE_EXPANSION/);

  const changed = initializeRepository();
  establishA1Baseline(changed);
  write(changed, 'packages/docx/src/renderer.ts', 'export function computePages() { return []; }\nexport function computeTableLayout() { return [1]; }\n');
  const changedResult = runChecker(changed, '--base-ref', 'main');
  assert.notEqual(changedResult.status, 0);
  assert.match(changedResult.output, /LEGACY_DECLARATION_CHANGED/);
});

test('allows only exact A2 service and option dependency threading through computePages', () => {
  const root = initializeRepository();
  establishA1Baseline(root);
  write(root, 'packages/docx/src/renderer.ts', 'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function createLayoutServices() {}\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n');

  const result = runChecker(root, '--base-ref', 'main');

  assert.equal(result.status, 0, result.output);
});

test('rejects unrelated computePages control-flow, calls, and parameters during A2 threading', () => {
  const cases = [
    'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return []; }\nexport function computeTableLayout() { return []; }\n',
    'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions) { buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n',
    'function buildMeasureState(ctx: unknown, fonts: unknown, services?: LayoutServices, options?: LayoutOptions) { return [ctx, fonts, services, options]; }\nexport function computePages(ctx: unknown, resolvedLocalFonts: unknown = {}, layoutServices?: LayoutServices, layoutOptions?: LayoutOptions, unrelated?: boolean) { const measure = buildMeasureState(ctx, resolvedLocalFonts, layoutServices, layoutOptions); return [measure]; }\nexport function computeTableLayout() { return []; }\n',
  ];
  for (const source of cases) {
    const root = initializeRepository();
    establishA1Baseline(root);
    write(root, 'packages/docx/src/renderer.ts', source);
    const result = runChecker(root, '--base-ref', 'main');
    assert.notEqual(result.status, 0);
    assert.match(result.output, /LEGACY_DECLARATION_CHANGED|BASELINE_EXPANSION/);
  }
});

test('final mode enforces the renderer adapter export and import allowlists', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-final-adapter-'));
  write(root, 'packages/docx/src/layout/document.ts', 'export function layoutDocument() {}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', 'export function paintLayoutPage() {}\n');
  write(root, 'packages/docx/src/renderer.ts', "import { layoutDocument } from './layout/document.js';\nimport { paintLayoutPage } from './paint/canvas-page.js';\nexport function paginateDocument() { return layoutDocument(); }\nexport function renderDocumentToCanvas() { return paintLayoutPage(); }\n");
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/renderer.ts', "import { hidden } from './hidden-layout.js';\nexport function paginateDocument() { return hidden(); }\nexport function renderDocumentToCanvas() {}\nexport function accidentalAlgorithm() {}\n");
  write(root, 'packages/docx/src/hidden-layout.ts', 'export function hidden() {}\n');
  const result = runChecker(root, '--final');
  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_ADAPTER_/);
});

test('final mode rejects layout logic hidden inside an allowed renderer adapter', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-inline-adapter-'));
  write(root, 'packages/docx/src/renderer.ts', `
export function paginateDocument(items: unknown[]) {
  const pages = [];
  for (const item of items) pages.push([item]);
  return pages;
}
export function renderDocumentToCanvas() {}
`);

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_ADAPTER_BODY/);
});

test('final mode rejects renamed fallback and style-cascade capabilities', () => {
  const diagnostic = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-diagnostic-fallback-'));
  write(diagnostic, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(diagnostic, 'packages/docx/src/layout/page.ts', "export const diagnosticFallback = { code: 'UNSUPPORTED_FEATURE' };\n");
  assert.equal(runChecker(diagnostic, '--final').status, 0);

  const fallback = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-old-engine-'));
  write(fallback, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(fallback, 'packages/docx/src/layout/page.ts', 'export const useOldEngine = true;\n');
  const fallbackResult = runChecker(fallback, '--final');
  assert.notEqual(fallbackResult.status, 0);
  assert.match(fallbackResult.output, /FINAL_LEGACY_BOUNDARY/);

  const style = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-fold-style-'));
  write(style, 'packages/docx/src/renderer.ts', 'export function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n');
  write(style, 'packages/docx/src/layout/page.ts', 'export function foldRunFormatting(base: object, direct: object) { return { ...base, ...direct }; }\n');
  const styleResult = runChecker(style, '--final');
  assert.notEqual(styleResult.status, 0);
  assert.match(styleResult.output, /LAYOUT_STYLE_CAPABILITY/);
});

test('final mode rejects star exports from renderer', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-star-export-'));
  write(root, 'packages/docx/src/layout/page.ts', 'export function accidentalAlgorithm() {}\n');
  write(root, 'packages/docx/src/renderer.ts', "export * from './layout/page.js';\nexport function paginateDocument() {}\nexport function renderDocumentToCanvas() {}\n");

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_ADAPTER_EXPORT/);
});

test('rejects rewriting a transitional baseline after A1', () => {
  const root = initializeRepository();
  establishA1Baseline(root);

  const result = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /TRANSITIONAL_BASELINE_EXISTS/);
});

test('rejects final mode while a transitional baseline remains', () => {
  const root = initializeRepository();
  establishA1Baseline(root);

  const result = runChecker(root, '--final', '--base-ref', 'main');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /FINAL_BASELINE_PRESENT/);
});
