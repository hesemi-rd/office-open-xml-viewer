import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const script = fileURLToPath(new URL('./check-docx-public-api.mjs', import.meta.url));

function git(root, ...args) {
  return execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
}

function fixture() {
  const root = mkdtempSync(path.join(tmpdir(), 'docx-public-api-'));
  const types = path.join(root, 'packages/docx/dist/types');
  mkdirSync(types, { recursive: true });
  writeFileSync(path.join(types, 'index.d.ts'), "export { Api } from './api.js';\n");
  writeFileSync(path.join(types, 'api.d.ts'), "import type { Detail } from './detail.js';\nexport declare class Api { detail: Detail; }\n");
  writeFileSync(
    path.join(types, 'detail.d.ts'),
    '/** private/sample-1 implementation evidence must not enter the public baseline. */\nexport interface Detail { value: string; }\n',
  );
  git(root, 'init', '-q');
  git(root, 'config', 'user.email', 'fixture@example.invalid');
  git(root, 'config', 'user.name', 'Fixture');
  git(root, 'add', '.');
  git(root, 'commit', '-qm', 'base declarations');
  return { root, base: git(root, 'rev-parse', 'HEAD') };
}

function run(root, ...args) {
  return spawnSync(process.execPath, [script, '--root', root, ...args], {
    cwd: root,
    encoding: 'utf8',
  });
}

test('writes a deterministic baseline containing every reachable local declaration', () => {
  const { root, base } = fixture();
  const result = run(root, '--base-ref', base, '--write-baseline');
  assert.equal(result.status, 0, result.stderr);
  const baseline = readFileSync(path.join(root, 'packages/docx/api/public-api-baseline.d.ts'), 'utf8');
  assert.match(baseline, /file: index\.d\.ts/);
  assert.match(baseline, /file: api\.d\.ts/);
  assert.match(baseline, /file: detail\.d\.ts/);
  assert.doesNotMatch(baseline, /private\/sample-1/);
  assert.equal(run(root, '--base-ref', base).status, 0);
});

test('fails when a transitively reachable declaration changes', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    path.join(root, 'packages/docx/dist/types/detail.d.ts'),
    'export interface Detail { value: number; }\n',
  );
  const result = run(root, '--base-ref', base);
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /public API declaration baseline differs/i);
});

test('ignores declaration-emitter quote style and redundant type parentheses', () => {
  const { root, base } = fixture();
  const detailPath = path.join(root, 'packages/docx/dist/types/detail.d.ts');
  writeFileSync(
    detailPath,
    'export interface A { a: string; }\nexport interface B { b: string; }\nexport type Detail = ({ kind: "a" } & A) | ({ kind: "b" } & B);\n',
  );
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  writeFileSync(
    detailPath,
    "export interface A { a: string; }\nexport interface B { b: string; }\nexport type Detail = { kind: 'a' } & A | { kind: 'b' } & B;\n",
  );

  const result = run(root, '--base-ref', base);

  assert.equal(result.status, 0, result.stderr);
});

test('fails when the generated entry declaration is missing', () => {
  const { root, base } = fixture();
  const result = run(root, '--base-ref', base, '--write-baseline', '--entry', 'missing.d.ts');
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /build.*DOCX package/i);
});

test('refuses to rewrite the baseline after it exists at the merge base', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  git(root, 'add', '.');
  git(root, 'commit', '-qm', 'establish baseline');
  const postA1 = git(root, 'rev-parse', 'HEAD');
  const result = run(root, '--base-ref', postA1, '--write-baseline');
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /only permitted before the merge base contains/i);
});

test('rejects manually changing both declarations and the committed baseline after A1', () => {
  const { root, base } = fixture();
  assert.equal(run(root, '--base-ref', base, '--write-baseline').status, 0);
  git(root, 'add', '.');
  git(root, 'commit', '-qm', 'establish baseline');
  const postA1 = git(root, 'rev-parse', 'HEAD');
  writeFileSync(
    path.join(root, 'packages/docx/dist/types/detail.d.ts'),
    'export interface Detail { value: number; }\n',
  );
  const baselinePath = path.join(root, 'packages/docx/api/public-api-baseline.d.ts');
  writeFileSync(baselinePath, readFileSync(baselinePath, 'utf8').replace('value: string', 'value: number'));

  const result = run(root, '--base-ref', postA1);

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /baseline differs from the merge base/i);
});
