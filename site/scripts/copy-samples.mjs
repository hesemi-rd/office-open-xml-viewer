// Copy the redistributable demo samples from each package into the site's
// public/ folder so the live showcase can fetch them. These office files are
// not committed under site/ (repo rule); they live in packages/*/public/demo.
import { copyFile, mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const out = resolve(here, '../public/samples');

const files = [
  ['packages/pptx/public/demo/sample-1.pptx', 'sample-1.pptx'],
  ['packages/xlsx/public/demo/sample-1.xlsx', 'sample-1.xlsx'],
  ['packages/docx/public/demo/sample-1.docx', 'sample-1.docx'],
];

await mkdir(out, { recursive: true });
for (const [src, name] of files) {
  await copyFile(resolve(root, src), resolve(out, name));
  console.log(`copied ${name}`);
}
