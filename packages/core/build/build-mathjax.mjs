// Pre-bundle the MathJax v4 + STIX Two Math converter into a single opaque
// asset (`assets/mathjax-stix2.js`). The renderer loads this asset lazily;
// keeping it pre-bundled (rather than importing @mathjax/src directly) stops the
// consuming app's bundler from re-bundling — and over-including — the MathJax
// source.
//
// The generated bundle is ~3 MB of minified IIFE, so — like the WASM parsers —
// it is NOT committed (see .gitignore) and is instead regenerated from source.
// This script runs automatically on install via core's `prepare` script; run it
// by hand with `pnpm --filter @silurus/ooxml-core build:mathjax`.
import { build } from 'esbuild';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = path.dirname(fileURLToPath(import.meta.url));

await build({
  entryPoints: [path.join(dir, 'stix2-entry.mjs')],
  bundle: true,
  minify: true,
  format: 'iife',
  target: 'es2020',
  legalComments: 'none',
  outfile: path.join(dir, '../assets/mathjax-stix2.js'),
});

console.log('built assets/mathjax-stix2.js');
