import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, statSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const coreDeclaration = path.join(root, 'packages/core/dist/types/index.d.ts');
const docxDeclaration = path.join(root, 'packages/docx/dist/types/index.d.ts');

function buildPackage(name) {
  execFileSync('pnpm', ['--filter', name, 'build'], {
    cwd: root,
    stdio: 'pipe',
  });
}

test('building DOCX does not rebuild declaration outputs owned by core', () => {
  buildPackage('@silurus/ooxml-core');
  const before = {
    contents: readFileSync(coreDeclaration),
    modifiedMs: statSync(coreDeclaration).mtimeMs,
  };

  buildPackage('@silurus/ooxml-docx');

  assert.equal(existsSync(docxDeclaration), true);
  assert.deepEqual(readFileSync(coreDeclaration), before.contents);
  assert.equal(statSync(coreDeclaration).mtimeMs, before.modifiedMs);
});
