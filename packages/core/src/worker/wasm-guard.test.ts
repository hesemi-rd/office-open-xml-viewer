import { describe, it, expect, vi } from 'vitest';
import { WasmParserHost, WasmTrapError, isWasmTrap } from './wasm-guard.js';

/** A stand-in for a `WebAssembly.RuntimeError` when the test runtime lacks a
 *  real WASM trap to throw. `WebAssembly.RuntimeError` exists in Node/jsdom, so
 *  we throw the genuine article where possible and fall back to this shape. */
function makeTrap(message = 'unreachable'): Error {
  const RE = (globalThis as { WebAssembly?: { RuntimeError?: new (m?: string) => Error } })
    .WebAssembly?.RuntimeError;
  if (RE) return new RE(message);
  const e = new Error(message);
  e.name = 'RuntimeError';
  return e;
}

describe('isWasmTrap', () => {
  it('flags a WebAssembly.RuntimeError (panic / unreachable / stack overflow)', () => {
    expect(isWasmTrap(makeTrap())).toBe(true);
  });

  it('flags a RangeError (failed memory.grow / OOM / stack overflow)', () => {
    expect(
      isWasmTrap(new RangeError('WebAssembly.Memory.grow(): Maximum memory size exceeded')),
    ).toBe(true);
    // V8 surfaces a native stack overflow as a RangeError too.
    expect(isWasmTrap(new RangeError('Maximum call stack size exceeded'))).toBe(true);
  });

  it('flags a trap-shaped NAME that crossed the worker boundary (prototype stripped)', () => {
    // structured-clone across the worker boundary drops the RuntimeError prototype
    // but keeps the name — match on name so a forwarded trap is still recognised.
    const runtime = new Error('unreachable');
    runtime.name = 'RuntimeError';
    expect(isWasmTrap(runtime)).toBe(true);
    const compile = new Error('bad module');
    compile.name = 'CompileError';
    expect(isWasmTrap(compile)).toBe(true);
    const link = new Error('import mismatch');
    link.name = 'LinkError';
    expect(isWasmTrap(link)).toBe(true);
  });

  it('does NOT flag a graceful parser error (Result::Err surfaced as a string)', () => {
    expect(isWasmTrap('pptx-parser error: not a zip archive')).toBe(false);
    expect(isWasmTrap(new Error('pptx-parser error: bad central directory'))).toBe(false);
    expect(isWasmTrap(new Error('serialize error: invalid utf-8'))).toBe(false);
  });

  it('does NOT flag a graceful "out of memory" wrapped in a plain Error (MINOR: no message-only trap detection)', () => {
    // A lenient degradation surfaced as `new Error('...out of memory...')` is a
    // GRACEFUL error (name === 'Error'), not a trap. Classifying it as a trap
    // would needlessly recycle a healthy instance and drop its archive. Only a
    // trap-shaped TYPE / NAME (RuntimeError / RangeError / CompileError / LinkError)
    // is a trap — never a substring on a plain Error.
    expect(isWasmTrap(new Error('xlsx-parser: sharedStrings too large, out of memory'))).toBe(
      false,
    );
    expect(isWasmTrap(new Error('memory access out of bounds'))).toBe(false);
    expect(isWasmTrap(new Error('recursion limit reached while parsing'))).toBe(false);
  });
});

describe('WasmParserHost', () => {
  it('runs a healthy operation and returns its value, no re-init', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const host = new WasmParserHost(init);
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    const out = host.run(() => 42);
    expect(out).toBe(42);
    expect(host.poisoned).toBe(false);
    // One init: the initial setWasmUrl. ensureReady did not re-init.
    expect(init).toHaveBeenCalledTimes(1);
  });

  it('passes a graceful error through unchanged and keeps the instance healthy', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const host = new WasmParserHost(init);
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    const graceful = new Error('pptx-parser error: not a zip');
    expect(() =>
      host.run(() => {
        throw graceful;
      }),
    ).toThrow(graceful);
    // Not a trap → not poisoned → no re-init on the next ready.
    expect(host.poisoned).toBe(false);
    await host.ensureReady();
    expect(init).toHaveBeenCalledTimes(1);
  });

  it('owns the archive handle and frees the prior one on setArchive', async () => {
    const freeArchive = vi.fn();
    const host = new WasmParserHost<{ id: string }>(vi.fn().mockResolvedValue(undefined), {
      freeArchive,
    });
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    const a1 = { id: 'a1' };
    host.setArchive(a1);
    expect(host.archive).toBe(a1);
    const a2 = { id: 'a2' };
    host.setArchive(a2); // frees a1 first
    expect(freeArchive).toHaveBeenCalledWith(a1);
    expect(host.archive).toBe(a2);
  });

  it('NEUTRALIZATION: a trap poisons the instance, frees + nulls the archive, and REINITS (not init) on the next request', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const reinit = vi.fn().mockResolvedValue(undefined);
    const freeArchive = vi.fn();
    const host = new WasmParserHost<{ id: string }>(init, { freeArchive, reinit });
    host.setWasmUrl('wasm://x');
    await host.ensureReady();

    // File #1: construct the archive inside the trapping closure (mirrors the
    // real parse path where `new PptxArchive(...)` then `.parse()` traps).
    const badArchive = { id: 'poisoned' };
    expect(() =>
      host.run(() => {
        host.setArchive(badArchive);
        throw makeTrap('unreachable');
      }),
    ).toThrow(WasmTrapError);
    // The instance is flagged dead, and the crashing handle was freed + nulled.
    expect(host.poisoned).toBe(true);
    expect(freeArchive).toHaveBeenCalledWith(badArchive);
    expect(host.archive).toBeNull();
    expect(init).toHaveBeenCalledTimes(1); // recovery is LAZY, not yet fired
    expect(reinit).toHaveBeenCalledTimes(0);

    // File #2 arrives: ensureReady rebuilds a FRESH instance via `reinit`, NOT by
    // re-running `init`. This is the load-bearing distinction: against the real
    // wasm-bindgen glue, re-`init` returns the cached poisoned instance and
    // recovers nothing, whereas `reinit` forces a genuine re-instantiation. So we
    // assert `init` was NOT called again and `reinit` WAS.
    await host.ensureReady();
    expect(init).toHaveBeenCalledTimes(1); // still just the first load
    expect(reinit).toHaveBeenCalledTimes(1); // the fresh instance came from reinit
    expect(reinit).toHaveBeenCalledWith('wasm://x');
    expect(host.poisoned).toBe(false);
    // ...and the next parse succeeds on clean linear memory.
    const good = host.run(() => {
      host.setArchive({ id: 'clean' });
      return 'parsed';
    });
    expect(good).toBe('parsed');
    expect(host.archive).toEqual({ id: 'clean' });
  });

  it('MUTATION GUARD: removing the reinit hook would recover via a no-op init — this test would then see init re-called', async () => {
    // Documents the mutation the reviewer flagged: if recovery fell back to
    // `init` (the wasm-bindgen singleton no-op), the assertion `init called once`
    // below fails. With `reinit` wired, `init` is never re-invoked.
    const init = vi.fn().mockResolvedValue(undefined);
    const reinit = vi.fn().mockResolvedValue(undefined);
    const host = new WasmParserHost(init, { reinit });
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    expect(() =>
      host.run(() => {
        throw makeTrap();
      }),
    ).toThrow(WasmTrapError);
    await host.ensureReady();
    // The single most important invariant: the recovery path did NOT re-run init.
    expect(init).toHaveBeenCalledTimes(1);
    expect(reinit).toHaveBeenCalledTimes(1);
  });

  it('does not double-free: after a trap frees the handle, the next parse frees only the NEW one', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const reinit = vi.fn().mockResolvedValue(undefined);
    const freed: string[] = [];
    const freeArchive = vi.fn((a: { id: string }) => freed.push(a.id));
    const host = new WasmParserHost<{ id: string }>(init, { freeArchive, reinit });
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    expect(() =>
      host.run(() => {
        host.setArchive({ id: 'first' });
        throw makeTrap();
      }),
    ).toThrow(WasmTrapError);
    await host.ensureReady();
    // The next parse's setArchive must NOT re-free 'first' (already freed on the
    // trap and nulled) — it frees nothing (host.archive is null), then adopts.
    host.setArchive({ id: 'second' });
    expect(freed).toEqual(['first']); // exactly once, no double-free
  });

  it('rebuilds from the SAME wasm url that was set originally', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const reinit = vi.fn().mockResolvedValue(undefined);
    const host = new WasmParserHost(init, { reinit });
    host.setWasmUrl('wasm://original');
    await host.ensureReady();
    expect(() =>
      host.run(() => {
        throw makeTrap();
      }),
    ).toThrow(WasmTrapError);
    await host.ensureReady();
    // First load through `init`, recovery through `reinit`, both with the same url.
    expect(init).toHaveBeenNthCalledWith(1, 'wasm://original');
    expect(reinit).toHaveBeenNthCalledWith(1, 'wasm://original');
  });

  it('falls back to init when no reinit hook is supplied (unit hosts only)', async () => {
    // Without `reinit` the host re-runs `init` — acceptable for unit tests, but a
    // no-op against the real singleton. Production wiring always passes `reinit`.
    const init = vi.fn().mockResolvedValue(undefined);
    const host = new WasmParserHost(init); // no reinit
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    expect(() =>
      host.run(() => {
        throw makeTrap();
      }),
    ).toThrow(WasmTrapError);
    await host.ensureReady();
    expect(init).toHaveBeenCalledTimes(2); // fallback re-ran init
  });

  it('stays poisoned if the rebuild itself fails, and retries on the following request', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const reinit = vi
      .fn()
      .mockRejectedValueOnce(new Error('network down')) // first respawn fails
      .mockResolvedValueOnce(undefined); // second respawn succeeds
    const host = new WasmParserHost(init, { reinit });
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    expect(() =>
      host.run(() => {
        throw makeTrap();
      }),
    ).toThrow(WasmTrapError);

    await expect(host.ensureReady()).rejects.toThrow('network down');
    expect(host.poisoned).toBe(true); // still dead

    await host.ensureReady(); // retries and recovers
    expect(host.poisoned).toBe(false);
    expect(init).toHaveBeenCalledTimes(1);
    expect(reinit).toHaveBeenCalledTimes(2);
  });

  it('a throwing freeArchive during poison does not mask the WasmTrapError', async () => {
    const init = vi.fn().mockResolvedValue(undefined);
    const freeArchive = vi.fn(() => {
      throw new Error('free() on poisoned memory');
    });
    const host = new WasmParserHost<{ id: string }>(init, { freeArchive });
    host.setWasmUrl('wasm://x');
    await host.ensureReady();
    expect(() =>
      host.run(() => {
        host.setArchive({ id: 'x' });
        throw makeTrap();
      }),
    ).toThrow(WasmTrapError);
    expect(host.poisoned).toBe(true);
    expect(host.archive).toBeNull();
  });
});

describe('WasmTrapError', () => {
  it('is an Error subclass carrying a stable code', () => {
    const err = new WasmTrapError('boom');
    expect(err).toBeInstanceOf(Error);
    expect(err).toBeInstanceOf(WasmTrapError);
    expect(err.code).toBe('parser-crashed');
    expect(err.name).toBe('WasmTrapError');
  });
});
