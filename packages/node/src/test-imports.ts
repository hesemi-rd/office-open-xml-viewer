/**
 * Test-only dynamic-import helpers.
 *
 * Several node test suites depend on modules that may be absent depending on
 * the build state or environment:
 *
 *   - `skia-canvas` — the native Canvas binding these canvas-backed checks use.
 *     It is a devDependency, so `pnpm install` provides it in CI as well as
 *     locally; but a contributor who has not installed it (or a stripped
 *     environment) should still be able to run the pure-logic suites.
 *   - the WASM-backed parser entrypoints (`./xlsx.ts`, `./docx.ts`, and the
 *     renderer sources they reach) — these statically import the git-ignored
 *     WASM glue, which only exists after `pnpm build:wasm`.
 *
 * Historically each suite loaded these via `await import(...).catch(() => null)`
 * and gated itself with `describe.skipIf(!mod)`. That is convenient locally, but
 * on CI it means a broken/missing dependency degrades to a SILENT skip: the job
 * stays green while testing nothing. To close that hole, CI sets
 * `OOXML_REQUIRE_SKIA=1`, which flips every such optional import from
 * "return null and skip" to "throw and fail the run". Locally the variable is
 * unset, so the suites keep skipping cleanly when a binding is absent.
 */

/** Truthy check for the CI opt-in flag (any non-empty value except "0"/"false"). */
function requireBindings(): boolean {
  const v = process.env.OOXML_REQUIRE_SKIA;
  return !!v && v !== '0' && v !== 'false';
}

/**
 * Run a dynamic `import()` for a test dependency.
 *
 * On success returns the module. On failure the behaviour depends on the
 * environment: when `OOXML_REQUIRE_SKIA` is truthy (CI) it re-throws an error
 * naming `what` and the underlying cause, turning a silent skip into a hard
 * failure; otherwise it returns `null` so the calling suite can
 * `describe.skipIf(!mod)` itself out.
 *
 * @param load - Thunk performing the `import()` (e.g. `() => import('skia-canvas')`).
 * @param what - Human-readable name of the dependency, used in the thrown error.
 */
export async function importForTests<T>(
  load: () => Promise<T>,
  what: string,
): Promise<T | null> {
  try {
    return await load();
  } catch (err) {
    if (requireBindings()) {
      const cause = err instanceof Error ? err.message : String(err);
      throw new Error(
        `OOXML_REQUIRE_SKIA is set but ${what} failed to load: ${cause}. ` +
          `In CI this dependency must be present (skia-canvas is a devDependency; ` +
          `the WASM parsers require \`pnpm build:wasm\` to have run first).`,
        { cause: err },
      );
    }
    return null;
  }
}

/**
 * Load `skia-canvas` for a test suite. Thin wrapper over {@link importForTests}
 * so every canvas-backed suite shares one gate: null → skip locally, throw →
 * fail under `OOXML_REQUIRE_SKIA=1` (CI).
 */
export function loadSkiaForTests(): Promise<typeof import('skia-canvas') | null> {
  return importForTests(() => import('skia-canvas'), 'skia-canvas');
}
