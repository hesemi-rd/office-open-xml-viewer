/**
 * Request/response correlation over a Web Worker.
 *
 * Each format package (pptx/docx/xlsx) drives a WASM parser in a Worker. The
 * naive pattern — attach a one-shot `message` listener keyed only on the
 * response `type` — breaks under concurrency: with two requests in flight, the
 * first arriving response of a matching type resolves the wrong promise (or, if
 * an unrelated message arrives, the listener detaches without resolving and the
 * promise hangs forever).
 *
 * `WorkerBridge` fixes this by assigning every request a monotonic id and
 * resolving against a pending-callback map keyed by that id — the proven
 * pattern. It is wire-protocol agnostic: the discriminant field (`kind` vs
 * `type`), the error shape, and any unsolicited messages (e.g. an init `ready`
 * handshake) are described by the {@link WorkerBridgeOptions} callbacks, so all
 * three packages can share one correlation mechanism without standardizing
 * their message envelopes.
 *
 * A correlated response is the happy path, but a worker can also fail *without*
 * ever posting one — a wedged parse loop, an uncaught script exception, a
 * structured-clone failure on the way back. Those leave requests hanging
 * forever. Three opt-in escape hatches settle them: a per-request `timeoutMs`,
 * an `AbortSignal`, and the worker's own `error` / `messageerror` events (which
 * reject *every* pending request, since the worker is presumed unusable). All
 * timers and listeners are torn down on settle so nothing leaks.
 */

/** The subset of the DOM `Worker` interface the bridge depends on. Keeping it
 *  structural lets tests substitute an in-memory fake. The real `Worker`
 *  satisfies this — it exposes the `error` / `messageerror` events too. */
export interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  addEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
  addEventListener(type: 'messageerror', listener: (e: MessageEvent) => void): void;
  addEventListener(type: 'error', listener: (e: ErrorEvent) => void): void;
  removeEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
  removeEventListener(type: 'messageerror', listener: (e: MessageEvent) => void): void;
  removeEventListener(type: 'error', listener: (e: ErrorEvent) => void): void;
  terminate(): void;
}

export interface WorkerBridgeOptions<TRes> {
  /**
   * Extract the correlation id from a response. Return `undefined` for
   * unsolicited messages that do not answer a request (e.g. a `ready`
   * handshake); those are routed to {@link onUnsolicited} instead.
   */
  readonly correlate: (response: TRes) => number | undefined;
  /**
   * Extract an error message from a response, or `undefined` when the response
   * is a success. When defined, the matching request rejects with `Error(msg)`.
   */
  readonly toError?: (response: TRes) => string | undefined;
  /** Called for every message that does not correlate to a pending request. */
  readonly onUnsolicited?: (response: TRes) => void;
  /**
   * Default per-request timeout in milliseconds. When set, every {@link
   * WorkerBridge.request} that does not override it rejects if the worker has
   * not answered within this window. Omit it (the default) to wait forever —
   * the historical behaviour. A per-request `timeoutMs` takes precedence.
   */
  readonly timeoutMs?: number;
}

/** Per-request escape-hatch options for {@link WorkerBridge.request}. */
export interface WorkerRequestOptions {
  /**
   * Reject this request if the worker has not answered within `timeoutMs`
   * milliseconds. Overrides the bridge-level default. Omit both to wait forever.
   */
  timeoutMs?: number;
  /**
   * Reject this request when the signal aborts (and immediately if it is
   * already aborted). The rejection is an `Error` whose `name` is
   * `'AbortError'`.
   */
  signal?: AbortSignal;
}

interface PendingEntry<TRes> {
  resolve: (r: TRes) => void;
  reject: (e: Error) => void;
  /** Tear down this request's timer and abort listener. Idempotent; called
   *  exactly once, on whichever path settles the request first. */
  cleanup: () => void;
}

/** Build an `AbortError`-flavoured Error without depending on the environment's
 *  `DOMException` constructor (absent in some worker/test runtimes). */
function abortError(): Error {
  const err = new Error('worker request aborted');
  err.name = 'AbortError';
  return err;
}

export class WorkerBridge<TRes = unknown> {
  private readonly _worker: WorkerLike;
  private readonly _opts: WorkerBridgeOptions<TRes>;
  private readonly _pending = new Map<number, PendingEntry<TRes>>();
  private _nextId = 1;

  constructor(worker: WorkerLike, opts: WorkerBridgeOptions<TRes>) {
    this._worker = worker;
    this._opts = opts;
    this._worker.addEventListener('message', this._handle);
    this._worker.addEventListener('messageerror', this._handleWorkerError);
    this._worker.addEventListener('error', this._handleWorkerError);
  }

  private _handle = (e: MessageEvent<TRes>): void => {
    const res = e.data;
    const id = this._opts.correlate(res);
    if (id === undefined) {
      this._opts.onUnsolicited?.(res);
      return;
    }
    const cb = this._pending.get(id);
    if (!cb) return; // unknown / already-settled id: ignore (never hang another request)
    this._pending.delete(id);
    cb.cleanup();
    const err = this._opts.toError?.(res);
    if (err !== undefined) cb.reject(new Error(err));
    else cb.resolve(res);
  };

  /**
   * The worker itself failed (uncaught script exception, load failure, or a
   * response that could not be deserialized). No correlated reply is coming, so
   * every in-flight request is rejected — the worker is presumed unusable.
   */
  private _handleWorkerError = (e: ErrorEvent | MessageEvent): void => {
    const detail = 'message' in e && e.message ? `: ${e.message}` : '';
    this._rejectAll(new Error(`Worker error${detail}`));
  };

  /** Reject and drain every pending request, tearing down each one's timer and
   *  abort listener. Shared by worker-error handling and {@link terminate}. */
  private _rejectAll(error: Error): void {
    const entries = [...this._pending.values()];
    this._pending.clear();
    for (const cb of entries) {
      cb.cleanup();
      cb.reject(error);
    }
  }

  /** Allocate the next correlation id. Useful when the caller must embed the id
   *  in a transferable-bearing message it builds itself. */
  nextId(): number {
    return this._nextId++;
  }

  /**
   * Send a correlated request and resolve with its matching response. `build`
   * receives the freshly allocated id so it can embed it in the message.
   *
   * `opts.timeoutMs` / `opts.signal` are opt-in escape hatches for a worker
   * that never replies; with neither (and no bridge-level default) the request
   * waits forever, as it always has.
   */
  request(
    build: (id: number) => unknown,
    transfer?: Transferable[],
    opts?: WorkerRequestOptions,
  ): Promise<TRes> {
    const id = this._nextId++;
    const timeoutMs = opts?.timeoutMs ?? this._opts.timeoutMs;
    const signal = opts?.signal;
    return new Promise<TRes>((resolve, reject) => {
      // Already-aborted signal: reject synchronously, register nothing.
      if (signal?.aborted) {
        reject(abortError());
        return;
      }

      let timer: ReturnType<typeof setTimeout> | undefined;
      let onAbort: (() => void) | undefined;

      // Runs once, on the first settling path (response, timeout, abort,
      // worker-error, or terminate), to release the timer and abort listener.
      const cleanup = (): void => {
        if (timer !== undefined) {
          clearTimeout(timer);
          timer = undefined;
        }
        if (onAbort && signal) {
          signal.removeEventListener('abort', onAbort);
          onAbort = undefined;
        }
      };

      this._pending.set(id, { resolve, reject, cleanup });

      if (timeoutMs !== undefined) {
        timer = setTimeout(() => {
          const cb = this._pending.get(id);
          if (!cb) return;
          this._pending.delete(id);
          cb.cleanup();
          cb.reject(new Error(`worker request timed out after ${timeoutMs}ms`));
        }, timeoutMs);
      }

      if (signal) {
        onAbort = (): void => {
          const cb = this._pending.get(id);
          if (!cb) return;
          this._pending.delete(id);
          cb.cleanup();
          cb.reject(abortError());
        };
        signal.addEventListener('abort', onAbort);
      }

      this._worker.postMessage(build(id), transfer);
    });
  }

  /** Fire-and-forget message with no correlation (e.g. the `init` message). */
  post(message: unknown, transfer?: Transferable[]): void {
    this._worker.postMessage(message, transfer);
  }

  /** Terminate the worker and reject every still-pending request. */
  terminate(): void {
    this._worker.removeEventListener('message', this._handle);
    this._worker.removeEventListener('messageerror', this._handleWorkerError);
    this._worker.removeEventListener('error', this._handleWorkerError);
    this._worker.terminate();
    this._rejectAll(new Error('Worker terminated'));
  }
}
