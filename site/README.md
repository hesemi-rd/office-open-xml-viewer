# @silurus/ooxml — showcase site

A standalone marketing / showcase site for [`@silurus/ooxml`](https://www.npmjs.com/package/@silurus/ooxml).
It embeds the **real** viewers (Rust/WASM + Canvas) rendering live sample
documents, alongside framework integration snippets.

This is separate from Storybook (`pnpm storybook`), which stays the test +
component workbench.

## Develop

```bash
pnpm build:wasm                      # build the parsers first (once)
pnpm --filter @silurus/ooxml-site dev
```

`predev` / `prebuild` copy the redistributable demo samples from
`packages/*/public/demo` into `public/samples/` (office files are not committed
under `site/`). The 3D hero asset `public/cube3.glb` is committed.

Dev-only `/preview/*` routes render individual sections in isolation (useful for
visual checks); they are not linked from the site.

## Build

```bash
pnpm --filter @silurus/ooxml-site build      # → site/dist/
SITE_BASE=/office-open-xml-viewer/ pnpm --filter @silurus/ooxml-site build  # project-pages base
```

## Deploy (GitHub Pages)

`.github/workflows/deploy-site.yml` builds and deploys to Pages, but is
**manual-only** (`workflow_dispatch`).

> ⚠️ GitHub Pages serves one site per repo. `deploy-storybook.yml` currently
> publishes Storybook to the custom domain `ooxml.silurus.dev`. Deploying this
> site **replaces** that. Decide which site owns Pages (and the custom domain)
> before enabling a push trigger.
