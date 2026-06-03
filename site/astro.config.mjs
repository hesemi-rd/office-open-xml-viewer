import { defineConfig } from 'astro/config';
import wasm from 'vite-plugin-wasm';
import topLevelAwait from 'vite-plugin-top-level-await';
import { fileURLToPath } from 'node:url';

const pkgSrc = (p) => fileURLToPath(new URL(`../packages/${p}/src/index.ts`, import.meta.url));

// GitHub Pages base path. Custom domain => '/', project pages => '/office-open-xml-viewer/'.
const SITE_BASE = process.env.SITE_BASE ?? '/';

// https://astro.build
export default defineConfig({
  base: SITE_BASE,
  trailingSlash: 'ignore',
  vite: {
    plugins: [wasm(), topLevelAwait()],
    worker: {
      format: 'es',
      plugins: () => [wasm(), topLevelAwait()],
    },
    resolve: {
      // Pull the workspace packages from source so Vite processes their
      // `?worker&inline` / `?url` imports (same flow as Storybook).
      alias: {
        '@silurus/ooxml-pptx': pkgSrc('pptx'),
        '@silurus/ooxml-xlsx': pkgSrc('xlsx'),
        '@silurus/ooxml-docx': pkgSrc('docx'),
        '@silurus/ooxml-core': pkgSrc('core'),
      },
    },
  },
});
