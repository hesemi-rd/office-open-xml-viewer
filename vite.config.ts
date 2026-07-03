import { defineConfig, type Plugin } from 'vite';
import wasm from 'vite-plugin-wasm';
import dts from 'vite-plugin-dts';
import { resolve, dirname, basename } from 'path';
import { fileURLToPath } from 'url';
import { readFile } from 'fs/promises';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Emit `?url` asset imports as real asset files instead of base64 `data:` URLs.
 * (Named hashlessly here — `rollupOptions.output.assetFileNames` is `[name]`.)
 *
 * Vite **library mode** force-inlines every `?url` asset as a
 * `data:<mime>;base64,…` string regardless of `assetsInlineLimit` (a number or a
 * `() => false` function does NOT override it — `build.lib` unconditionally
 * returns `true` from Vite's internal `shouldInline`). Two heavy asset kinds ride
 * on this path:
 *   - the three parser WASM modules (`*_parser_bg.wasm?url`, ~0.6–0.7 MB each) —
 *     base64 inflates them +33 % and blocks `WebAssembly.compileStreaming`
 *     (a data URL cannot be fetch-streamed; the worker must `atob` by hand);
 *   - the MathJax + STIX Two Math engine (`assets/mathjax-stix2.js?url`, ~3 MB) —
 *     inlined it turned the opt-in `math.mjs` chunk into a 4.1 MB base64 blob,
 *     even though consumers only import it when a document actually has equations.
 *
 * All of these are `?url` imports in a single owner module each (the format
 * main-thread handles — `document.ts` / `presentation.ts` / `workbook.ts` — and
 * `math/engine.ts`). We intercept the `?url` variant here, `emitFile` the bytes
 * as an asset next to the chunk, and hand back the standard ESM asset reference
 * `new URL('<name>', import.meta.url)` — the form Vite / webpack 5 / Rollup /
 * esbuild all rewrite when they re-bundle our `.mjs`, and which resolves
 * correctly for a plain `<script type=module>` too. wasm-bindgen's `--target web`
 * glue then `fetch()`es its URL and hits `instantiateStreaming`; the math engine
 * is lazy-loaded via a `<script src>` pointed at the emitted asset.
 *
 * Runs with `enforce: 'pre'` and claims every `?url` import; a bare-`.wasm`
 * import (owned by `vite-plugin-wasm`) is untouched. Nothing in the tree imports
 * bare `.wasm`, and every `?url` here is a real on-disk asset we want emitted —
 * exactly what Vite's non-lib mode would do anyway.
 */
function wasmAssetUrl(): Plugin {
  const SUFFIX = '?url';
  return {
    name: 'wasm-asset-url',
    enforce: 'pre',
    // Build-only: emitFile/ROLLUP_FILE_URL are Rollup build machinery and do
    // not exist on the dev server. In dev, Vite's stock `?url` handling serves
    // the file directly (the pre-E4 behavior) — intercepting there returned an
    // unresolvable reference and broke every WASM load (caught by CI smoke).
    apply: 'build',
    async load(id) {
      if (!id.endsWith(SUFFIX)) return null;
      const filePath = id.slice(0, -SUFFIX.length);
      const source = await readFile(filePath);
      const referenceId = this.emitFile({
        type: 'asset',
        name: basename(filePath),
        source,
      });
      // `import.meta.ROLLUP_FILE_URL_<id>` expands to the emitted asset's URL at
      // render time; wrapping it in `new URL(…, import.meta.url)` yields an
      // absolute href the worker (or the math engine's `<script>` loader) can
      // fetch from any realm.
      return `export default new URL(import.meta.ROLLUP_FILE_URL_${referenceId}, import.meta.url).href;`;
    },
  };
}

export default defineConfig({
  plugins: [
    wasmAssetUrl(),
    wasm(),
    dts({
      include: [
        'src/**/*',
        'packages/core/src/**/*',
        'packages/pptx/src/**/*',
        'packages/xlsx/src/**/*',
        'packages/docx/src/**/*',
      ],
      outDir: 'dist/types',
      tsconfigPath: './tsconfig.lib.json',
      rollupTypes: true,
      skipDiagnostics: true,
    }),
  ],
  build: {
    lib: {
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        pptx:  resolve(__dirname, 'src/pptx.ts'),
        xlsx:  resolve(__dirname, 'src/xlsx.ts'),
        docx:  resolve(__dirname, 'src/docx.ts'),
        // Opt-in math engine (MathJax + STIX Two Math). Separate entry so the
        // ~3 MB asset stays out of the docx/pptx bundles unless imported.
        math:  resolve(__dirname, 'src/math.ts'),
      },
      // ESM-only: the published bundle inlines a large math engine; emitting a
      // duplicate CJS copy of every chunk roughly doubled the package size.
      // Every modern bundler (Vite / webpack / Rollup / esbuild / Next) and
      // Node ≥ 20 consume ESM, so we ship `.mjs` only.
      formats: ['es'],
      fileName: (_format, name) => `${name}.mjs`,
    },
    rollupOptions: {
      output: { assetFileNames: '[name][extname]' },
    },
    target: 'esnext',
  },
  worker: {
    format: 'es',
    plugins: () => [wasm()],
  },
});
