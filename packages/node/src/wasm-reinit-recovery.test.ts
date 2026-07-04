import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
// Deep-import core by relative path (the node package doesn't list core as a
// dependency; sibling tests reach into `../../core/src/...` the same way).
import { WasmParserHost, isWasmTrap } from '../../core/src/worker/wasm-guard.ts';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — wasm-pack generated JS without a d.ts entry for the bare module path
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { resolveWasm } from './wasm-loader.ts';

/**
 * RB6 CRITICAL proof (against the REAL wasm-bindgen glue, not a mock).
 *
 * The bug the review found: wasm-bindgen keeps its instance in a module-level
 * singleton and `init` short-circuits (`if (wasm !== undefined) return wasm;`) on
 * every later call. So `WasmParserHost.ensureReady()` re-running `init` after a
 * trap recovers NOTHING — the poisoned instance and its corrupt linear memory are
 * handed straight back. A unit test that only counts `init` calls can't see this.
 *
 * These tests exercise the fix — the appended `reinit` export
 * (`scripts/append-wasm-reinit.mjs`) plumbed through `WasmParserHost` — on real
 * memory, and assert the two facts a mock cannot fake:
 *   1. `init` re-run is a NO-OP: a sentinel byte written to the instance's linear
 *      memory SURVIVES a second `init` (same buffer object) — proving the bug.
 *   2. `reinit` is a genuine re-instantiation: after `reinit`, the memory buffer
 *      is a DIFFERENT object and the sentinel is GONE — proving recovery works.
 *
 * A `WasmParserHost` wired with the real `reinit` is then driven through a
 * simulated trap → recovery cycle to confirm the fresh instance surfaces end to
 * end.
 */

interface GlueModule {
  default: (input: { module_or_path: WebAssembly.Module }) => Promise<{ memory: WebAssembly.Memory }>;
  reinit: (input: { module_or_path: WebAssembly.Module }) => Promise<{ memory: WebAssembly.Memory }>;
}

const FORMATS: ReadonlyArray<{ name: string; glue: unknown; wasmRel: string }> = [
  { name: 'pptx', glue: pptxWasm, wasmRel: '../../pptx/src/wasm/pptx_parser_bg.wasm' },
  { name: 'docx', glue: docxWasm, wasmRel: '../../docx/src/wasm/docx_parser_bg.wasm' },
  { name: 'xlsx', glue: xlsxWasm, wasmRel: '../../xlsx/src/wasm/xlsx_parser_bg.wasm' },
];

/** Write a sentinel near the top of the current linear memory and read it back to
 *  confirm the write landed. Returns the offset + sentinel so a later read can
 *  check whether that exact instance's memory is still around. */
function stampSentinel(memory: WebAssembly.Memory): { offset: number; value: number } {
  const view = new Uint8Array(memory.buffer);
  const offset = view.length - 16;
  const value = 0xa5;
  view[offset] = value;
  expect(view[offset]).toBe(value); // the write took on THIS buffer
  return { offset, value };
}

describe.each(FORMATS)('RB6 real-glue recovery: $name parser', ({ glue, wasmRel }) => {
  const mod = glue as GlueModule;
  const compiled = new WebAssembly.Module(readFileSync(resolveWasm(import.meta.url, wasmRel)));

  it('re-running init is a NO-OP: the poisoned memory survives (this is the bug the fix targets)', async () => {
    const exports1 = await mod.default({ module_or_path: compiled });
    const { offset, value } = stampSentinel(exports1.memory);

    // Re-run init — wasm-bindgen returns the CACHED instance, so nothing is rebuilt.
    const exports2 = await mod.default({ module_or_path: compiled });

    // Same instance, same memory buffer, sentinel intact: init "recovery" is a lie.
    expect(exports2.memory.buffer).toBe(exports1.memory.buffer);
    expect(new Uint8Array(exports2.memory.buffer)[offset]).toBe(value);
  });

  it('reinit is a REAL re-instantiation: fresh memory buffer, the sentinel is GONE', async () => {
    const exports1 = await mod.default({ module_or_path: compiled });
    const { offset, value } = stampSentinel(exports1.memory);

    // reinit nulls the singleton first, forcing a genuine WebAssembly.instantiate.
    const exports2 = await mod.reinit({ module_or_path: compiled });

    // A DIFFERENT WebAssembly.Memory: the corrupt heap was discarded.
    expect(exports2.memory).not.toBe(exports1.memory);
    expect(exports2.memory.buffer).not.toBe(exports1.memory.buffer);
    // The state written to the (now-freed) old instance is GONE on the fresh one.
    expect(new Uint8Array(exports2.memory.buffer)[offset]).not.toBe(value);
    expect(new Uint8Array(exports2.memory.buffer)[offset]).toBe(0);
  });

  it('WasmParserHost drives a trap → recovery cycle onto a FRESH instance (MUTATION: init-only recovery would fail here)', async () => {
    // Wire the host with the REAL glue `init` and `reinit`. This is the true
    // discriminator: if the host's recovery path used `init` (the wasm-bindgen
    // singleton no-op) instead of `reinit`, the poisoned memory would be handed
    // back and the sentinel below would SURVIVE, failing the final assertion.
    // Track how the live instance moves by reading the module's exported memory
    // through the glue each time.
    const host = new WasmParserHost<unknown>(
      (input) => mod.default(input as unknown as { module_or_path: WebAssembly.Module }),
      {
        reinit: (input) => mod.reinit(input as unknown as { module_or_path: WebAssembly.Module }),
      },
    );
    // Ensure a known baseline: force a fresh instance before this test so the
    // singleton isn't whatever a sibling test left cached, then let the host load.
    const baseline = await mod.reinit({ module_or_path: compiled });
    const sentinel = stampSentinel(baseline.memory);
    host.setWasmUrl({ module_or_path: compiled } as unknown as string);
    await host.ensureReady(); // host's `init` returns the cached (stamped) instance

    // Sanity: the host is running on the SAME stamped instance right now (init is
    // the singleton), so the sentinel is still present pre-recovery.
    const memBeforeTrap = (await mod.default({ module_or_path: compiled })).memory;
    expect(memBeforeTrap).toBe(baseline.memory);
    expect(new Uint8Array(memBeforeTrap.buffer)[sentinel.offset]).toBe(sentinel.value);

    // File #1 traps. `isWasmTrap` must classify a real RuntimeError as a trap so
    // the host poisons + schedules recovery.
    const trap = new WebAssembly.RuntimeError('unreachable');
    expect(isWasmTrap(trap)).toBe(true);
    expect(() =>
      host.run(() => {
        throw trap;
      }),
    ).toThrow(/parser trapped and was recycled/);
    expect(host.poisoned).toBe(true);

    // File #2: ensureReady rebuilds via `reinit` → a genuinely different instance.
    await host.ensureReady();
    expect(host.poisoned).toBe(false);
    // Read the module's memory again: after `reinit` it is a DIFFERENT
    // WebAssembly.Memory and the pre-trap sentinel is GONE. If recovery had used
    // `init`, this would still be `baseline.memory` with the sentinel intact.
    const memAfterRecovery = (await mod.default({ module_or_path: compiled })).memory;
    expect(memAfterRecovery).not.toBe(baseline.memory);
    expect(new Uint8Array(memAfterRecovery.buffer)[sentinel.offset]).not.toBe(sentinel.value);
  });
});
