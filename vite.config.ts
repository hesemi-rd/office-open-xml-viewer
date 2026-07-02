import { defineConfig, type Plugin } from 'vite';
import wasm from 'vite-plugin-wasm';
import dts from 'vite-plugin-dts';
import { resolve, dirname, basename } from 'path';
import { fileURLToPath } from 'url';
import { readFile } from 'fs/promises';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Emit `*.wasm?url` imports as real asset files instead of base64 `data:` URLs.
 *
 * Vite **library mode** force-inlines every `?url` asset as a
 * `data:application/wasm;base64,…` string regardless of `assetsInlineLimit`
 * (a number or a `() => false` function does NOT override it). That inflates the
 * WASM by +33 % and blocks `WebAssembly.compileStreaming` (a data URL cannot be
 * fetch-streamed — the worker has to `atob` the base64 by hand).
 *
 * The single `.wasm?url` import lives in each format's main-thread handle
 * (`document.ts` / `presentation.ts` / `workbook.ts`). We intercept it here,
 * `emitFile` the bytes as a hashless asset next to the chunk, and hand back the
 * standard ESM asset reference `new URL('<name>.wasm', import.meta.url)` — the
 * form Vite / webpack 5 / Rollup / esbuild all rewrite when they re-bundle our
 * `.mjs`, and which resolves correctly for a plain `<script type=module>` too.
 * wasm-bindgen's `--target web` glue then `fetch()`es that URL and hits
 * `instantiateStreaming`.
 *
 * Runs with `enforce: 'pre'` and only claims the `?url` variant; the bare-`.wasm`
 * import (owned by `vite-plugin-wasm`) is untouched — though nothing in the tree
 * imports bare `.wasm`, so in practice this is the only WASM interception point.
 */
function wasmAssetUrl(): Plugin {
  const SUFFIX = '.wasm?url';
  return {
    name: 'wasm-asset-url',
    enforce: 'pre',
    async load(id) {
      if (!id.endsWith(SUFFIX)) return null;
      const filePath = id.slice(0, -'?url'.length);
      const source = await readFile(filePath);
      const referenceId = this.emitFile({
        type: 'asset',
        name: basename(filePath),
        source,
      });
      // `import.meta.ROLLUP_FILE_URL_<id>` expands to the emitted asset's URL at
      // render time; wrapping it in `new URL(…, import.meta.url)` yields an
      // absolute href the worker can fetch from any realm.
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
