/**
 * Self-poison + auto-respawn for a Rust/WASM parser instance (RB6).
 *
 * The three format parsers (pptx / docx / xlsx) are compiled with
 * `panic = "abort"` (see the workspace `Cargo.toml`): a Rust panic, an
 * `unreachable`, an out-of-memory `memory.grow`, or a stack overflow does not
 * unwind — it *traps*. A WASM trap leaves the module instance's linear memory in
 * an indeterminate state, so **every subsequent exported call on that same
 * instance is unsafe** (it may trap again, or silently return garbage from a
 * half-mutated heap).
 *
 * Each format package runs its parser in one long-lived worker that reuses a
 * single WASM instance across many `parse` / `extractImage` / … calls. Without
 * this guard, the first file that traps poisons the instance and takes every
 * *later* file down with it — one malicious or malformed document kills the
 * whole viewer session. That is exactly the RB6 threat.
 *
 * {@link WasmParserHost} draws the line between the two failure modes a parser
 * call can produce:
 *
 *   - A **graceful error** — the Rust code returned `Result::Err`, surfaced by
 *     wasm-bindgen as a thrown JS string/Error. The instance is *healthy*; the
 *     file was simply unparseable. Pass it through unchanged, keep the instance.
 *   - A **trap** — a `WebAssembly.RuntimeError` (panic / unreachable / stack
 *     overflow) or a `RangeError` (a failed `memory.grow`). The instance is
 *     *poisoned*. Mark it dead, drop the archive handle, and re-init a fresh
 *     module on the next call so the following file parses on clean memory.
 *
 * The recovery is intentionally lazy: {@link WasmParserHost.poison} only flags
 * the instance, and the fresh `init()` runs inside the next {@link
 * WasmParserHost.ensureReady}. A worker that traps and is never used again pays
 * nothing; a worker that keeps going gets a clean instance exactly when it needs
 * one. This lives in `core` so all three formats share one implementation (the
 * only per-format pieces are the `init` function and the archive type, injected
 * by the caller).
 */

/** Machine-readable code for the RB6 self-heal error. */
export type WasmTrapErrorCode = 'parser-crashed';

/**
 * Typed error thrown by {@link WasmParserHost.run} when the guarded WASM call
 * *trapped* (as opposed to returning a graceful `Result::Err`). Signals that
 * this one file crashed the parser and the instance was recycled — the caller
 * may keep using the same engine for the next file.
 *
 * Like {@link import('../errors/ooxml-error').OoxmlError}, the `instanceof`
 * check does not survive a structured-clone across the worker boundary, so
 * workers forward the `code` string and the main thread reconstructs the error.
 */
export class WasmTrapError extends Error {
  readonly code: WasmTrapErrorCode = 'parser-crashed';

  constructor(message: string) {
    super(message);
    this.name = 'WasmTrapError';
    // Restore the prototype chain for down-levelled `extends Error` targets so
    // `instanceof WasmTrapError` holds.
    Object.setPrototypeOf(this, WasmTrapError.prototype);
  }
}

/**
 * Whether `err` indicates the WASM instance is *poisoned* (its linear memory is
 * no longer trustworthy), as opposed to a graceful parser error.
 *
 * A `panic = "abort"` build turns a Rust panic / `unreachable` / stack overflow
 * into a `WebAssembly.RuntimeError`, and an out-of-memory `memory.grow` into a
 * `RangeError` ("WebAssembly.Memory.grow" / "out of memory" / "Out of memory").
 * Both leave the instance unusable. A `Result::Err(JsValue::from_str(...))`
 * arrives as a thrown *string* or a plain `Error`, which this returns `false`
 * for — the instance survives that.
 *
 * The discriminator is the error's TYPE / NAME, never its message substring
 * alone. A genuine trap always carries a trap-shaped constructor: a
 * `WebAssembly.RuntimeError` / `CompileError` / `LinkError` (matched by
 * `instanceof` and by `name`, since the `instanceof` check does not survive a
 * structured-clone across the worker boundary), or a `RangeError` for a failed
 * allocation. A plain `Error` (name `'Error'`) is a GRACEFUL parser error even
 * when its message mentions "out of memory" — e.g. a wrapper that surfaces a
 * lenient degradation as `new Error('...out of memory...')`. Classifying that as
 * a trap would needlessly recycle a healthy instance and drop its open archive,
 * so message-substring sniffing is deliberately NOT done here.
 */
export function isWasmTrap(err: unknown): boolean {
  // `WebAssembly.RuntimeError` is the canonical trap. Guard the reference in
  // case a runtime lacks the constructor (older / non-WASM test envs).
  const RuntimeError = (globalThis as { WebAssembly?: { RuntimeError?: unknown } }).WebAssembly
    ?.RuntimeError as (new () => Error) | undefined;
  if (RuntimeError && err instanceof RuntimeError) return true;
  // A failed allocation (`memory.grow` past the maximum) surfaces as a
  // RangeError rather than a RuntimeError; the instance is equally unusable.
  if (err instanceof RangeError) return true;
  // A trap that crossed the worker boundary (structured-clone strips the
  // prototype) or was re-thrown as a generic Error still carries the trap-shaped
  // NAME. Match those, but require the name — never a message substring on a
  // plain `Error`, which would over-match a graceful "out of memory" parser
  // error and poison a healthy instance.
  if (err instanceof Error) {
    const name = err.name;
    if (name === 'RuntimeError' || name === 'CompileError' || name === 'LinkError') return true;
  }
  return false;
}

/**
 * The input a wasm-bindgen `init(input)` accepts: a URL string / `URL` to fetch,
 * or the already-decoded module bytes (a `data:` URL that `init` cannot fetch is
 * decoded to an `ArrayBuffer`/`BufferSource` by the worker first). Kept wide to
 * match the generated glue's `InitInput` without depending on its exact type.
 */
export type WasmInitInput = string | URL | ArrayBuffer | ArrayBufferView | Response;

/** How a {@link WasmParserHost} initialises its WASM module. Mirrors the
 *  wasm-bindgen `init(input)` free function each parser package exports. */
export type WasmInit = (input: WasmInitInput) => Promise<unknown>;

/**
 * How a {@link WasmParserHost} REBUILDS its WASM module after a trap. Mirrors the
 * `reinit(input)` free function `scripts/append-wasm-reinit.mjs` appends to each
 * parser glue.
 *
 * This is distinct from {@link WasmInit} for a load-bearing reason: wasm-bindgen's
 * generated `init` keeps its instance in a module-level singleton and
 * short-circuits (`if (wasm !== undefined) return wasm;`) on every later call, so
 * re-running `init` after a trap hands back the SAME poisoned instance and
 * "recovery" silently does nothing. `reinit` nulls that singleton first, forcing a
 * genuine `WebAssembly.instantiate` with fresh linear memory. The host therefore
 * MUST use `reinit` (not `init`) on the recovery path — see
 * {@link WasmParserHost.ensureReady}.
 */
export type WasmReinit = (input: WasmInitInput) => Promise<unknown>;

/** Optional per-host hooks. */
export interface WasmParserHostOptions<TArchive> {
  /**
   * How to free an archive handle (e.g. `(a) => a.free()`). Called when the host
   * replaces the handle on a re-parse, and once on a trap right before the
   * instance is recycled. A throwing `free()` on a trapped handle is caught and
   * swallowed (see {@link WasmParserHost.run}). Omit for a host that keeps no
   * archive.
   */
  readonly freeArchive?: (archive: TArchive) => void;
  /**
   * How to force a FRESH instance after a trap (the glue's `reinit` export). If
   * omitted, the host falls back to calling {@link WasmInit} again — but note that
   * against the real wasm-bindgen singleton that fallback is a NO-OP (the poisoned
   * instance is reused). Production callers MUST pass this; it is optional only so
   * unit tests that assert lifecycle bookkeeping can omit it. When present it is
   * the sole re-instantiation path; `init` is used only for the very first load.
   */
  readonly reinit?: WasmReinit;
}

/**
 * Owns the lifecycle of one WASM parser instance **and its archive handle**, and
 * recycles both after a trap.
 *
 * Making the host the single owner of the archive is what makes recovery safe:
 * on a trap it frees the handle and nulls its own reference in one place, so the
 * worker never double-frees a handle that belonged to a now-discarded instance.
 * The worker reads the current handle through {@link WasmParserHost.archive}
 * instead of a local variable.
 *
 * Usage inside a worker:
 *
 * ```ts
 * const host = new WasmParserHost<PptxArchive>(init, { freeArchive: (a) => a.free() });
 * // on the init message:
 * host.setWasmUrl(wasmUrl);
 * // on a parse message:
 * await host.ensureReady();               // re-inits transparently if poisoned
 * const model = host.run(() => {          // traps → free handle + WasmTrapError
 *   host.setArchive(new PptxArchive(bytes, max)); // frees any prior handle first
 *   return host.archive!.parse();
 * });
 * // on a later message:
 * const bytes = host.run(() => host.archive!.extract_image(path));
 * ```
 *
 * `run` is synchronous because a wasm-bindgen call is synchronous; the async
 * work (a fresh `init`) is confined to `ensureReady`, which the caller already
 * awaits before every request.
 */
export class WasmParserHost<TArchive = unknown> {
  private readonly _init: WasmInit;
  private readonly _opts: WasmParserHostOptions<TArchive>;
  private _wasmInput: WasmInitInput | null = null;
  private _initPromise: Promise<unknown> | null = null;
  /** Set true by {@link poison} after a trap; cleared by the next re-init. */
  private _poisoned = false;
  /** The live archive handle for the current instance, or null. Owned here so a
   *  trap frees + nulls it in exactly one place (no worker-side double-free). */
  private _archive: TArchive | null = null;

  constructor(init: WasmInit, opts: WasmParserHostOptions<TArchive> = {}) {
    this._init = init;
    this._opts = opts;
  }

  /** Record the WASM URL/input and kick off the first `init`. Called once, on
   *  the worker's `init` message — same moment the old code did `init(url)`. */
  setWasmUrl(input: WasmInitInput): void {
    this._wasmInput = input;
    this._poisoned = false;
    this._initPromise = this._init(input);
  }

  /** The archive handle for the current instance, or null when none is open
   *  (before the first parse, or after a trap recycled the instance). */
  get archive(): TArchive | null {
    return this._archive;
  }

  /**
   * Adopt a freshly constructed archive handle as the host's own, freeing any
   * prior handle first (the re-parse dispose the workers used to do inline).
   */
  setArchive(archive: TArchive): void {
    this._freeArchive();
    this._archive = archive;
  }

  /** Free the current handle (if any) and null it — the double-free / UAF guard
   *  the workers used to keep as a local `disposeArchive()`. */
  disposeArchive(): void {
    this._freeArchive();
  }

  private _freeArchive(): void {
    if (this._archive != null && this._opts.freeArchive) {
      this._opts.freeArchive(this._archive);
    }
    this._archive = null;
  }

  /** True after a trap and before the next successful re-init. Exposed for
   *  tests and diagnostics. */
  get poisoned(): boolean {
    return this._poisoned;
  }

  /**
   * Resolve when a *healthy* instance is ready. If the instance was poisoned,
   * this transparently re-inits a fresh module from the retained URL first, so
   * the next {@link run} executes on clean linear memory.
   *
   * Mirrors the historical `await initPromise` the workers already do before
   * each request — callers just swap that one line for `await ensureReady()`.
   */
  async ensureReady(): Promise<void> {
    if (this._poisoned) {
      if (this._wasmInput === null) {
        throw new Error('WasmParserHost: setWasmUrl was never called');
      }
      // Recovery MUST use `reinit`, not `init`: wasm-bindgen's `init` returns the
      // cached (poisoned) instance on every later call, so re-`init`-ing recovers
      // nothing. `reinit` nulls the singleton first, forcing a real
      // `WebAssembly.instantiate` with fresh linear memory. Fall back to `init`
      // only when no `reinit` was supplied (unit tests that don't drive real
      // glue); production callers always pass `reinit`.
      const rebuild = this._opts.reinit ?? this._init;
      // Only flip `_poisoned` off once the rebuild settles successfully, so a
      // failed rebuild stays poisoned and is retried on the following request
      // rather than handing back a dead module.
      const p = rebuild(this._wasmInput);
      this._initPromise = p;
      await p;
      this._poisoned = false;
      return;
    }
    if (this._initPromise === null) {
      throw new Error('WasmParserHost: setWasmUrl was never called');
    }
    await this._initPromise;
  }

  /**
   * Run a synchronous WASM operation. A graceful error propagates unchanged
   * (the instance stays healthy). A *trap* ({@link isWasmTrap}) poisons the
   * instance — freeing + nulling {@link archive} and scheduling a re-init on the
   * next {@link ensureReady} — and is rethrown as a {@link WasmTrapError} so the
   * caller can report "this one file crashed the parser" while the next file
   * recovers on a fresh instance.
   */
  run<T>(op: () => T): T {
    try {
      return op();
    } catch (err) {
      if (isWasmTrap(err)) {
        this._poison();
        const detail = err instanceof Error ? err.message : String(err);
        throw new WasmTrapError(`WASM parser trapped and was recycled: ${detail}`);
      }
      throw err;
    }
  }

  /**
   * Mark the instance poisoned without running an operation — for the rare case
   * where the caller detects corruption out of band. The next
   * {@link ensureReady} re-inits.
   */
  poison(): void {
    this._poison();
  }

  private _poison(): void {
    this._poisoned = true;
    // Drop the init promise so ensureReady takes the re-init branch even if the
    // poison flag were ever cleared out of order.
    this._initPromise = null;
    // Free the archive that belongs to the now-dead instance and null the host's
    // sole reference, so the next parse's setArchive/disposeArchive never
    // double-frees a handle from a discarded module. A `free()` on a trapped
    // handle may itself throw (its memory is poisoned) — swallow so it never
    // masks the original trap.
    if (this._archive != null && this._opts.freeArchive) {
      try {
        this._opts.freeArchive(this._archive);
      } catch {
        // expected on poisoned memory; ignore
      }
    }
    this._archive = null;
  }
}
