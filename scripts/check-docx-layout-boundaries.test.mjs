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
  write(root, 'packages/docx/src/renderer.ts', 'export function computePages() { return []; }\n');
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
  assert.match(result.output, /TRANSITIVE_FORBIDDEN_EDGE/);
  assert.match(result.output, /canvas-page\.ts.*helper\.ts.*line-layout\.ts/s);
});

test('rejects a CommonJS require edge that bypasses static ESM imports', () => {
  const root = mkdtempSync(join(tmpdir(), 'docx-layout-boundary-require-'));
  write(root, 'packages/docx/src/renderer.ts', 'export const adapter = true;\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts', "const helper = require('./helper.js');\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/helper.ts', "const measured = require('../line-layout.js');\nexport { measured };\n");
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() { return []; }\n');

  const result = runChecker(root, '--final');

  assert.notEqual(result.status, 0);
  assert.match(result.output, /TRANSITIVE_FORBIDDEN_EDGE/);
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
  assert.match(result.output, /BASELINE_MISMATCH/);
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
