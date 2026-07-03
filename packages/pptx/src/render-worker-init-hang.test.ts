import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

/**
 * AR4 (worker-render mode twin of worker-init-hang): the render-capable worker
 * carried the same `ready`-flag hazard. After the initPromise conversion a
 * REJECTED WASM init must reject the pending request (`error` response) rather
 * than leaving `load()` hanging. Driven against a mocked WASM module + stubbed
 * `self`; only the parse arm is exercised (no OffscreenCanvas render needed).
 */

const initMock = vi.fn();
class FakePptxArchive {
  constructor(_bytes: Uint8Array, _max?: bigint) {}
  parse(): Uint8Array {
    return new TextEncoder().encode('{"slides":[],"slideWidth":0,"slideHeight":0}');
  }
  extract_media(): Uint8Array {
    return new Uint8Array([1]);
  }
  extract_image(): Uint8Array {
    return new Uint8Array([2]);
  }
  free(): void {}
}

vi.mock('./wasm/pptx_parser.js', () => ({
  default: (arg: unknown) => initMock(arg),
  PptxArchive: FakePptxArchive,
}));

interface FakeSelf {
  onmessage: ((e: MessageEvent) => void) | null;
  posted: unknown[];
  postMessage: (msg: unknown, transfer?: Transferable[]) => void;
}

function installSelf(): FakeSelf {
  const posted: unknown[] = [];
  const fake: FakeSelf = {
    onmessage: null,
    posted,
    postMessage: (msg: unknown) => {
      posted.push(msg);
    },
  };
  vi.stubGlobal('self', fake);
  return fake;
}

async function loadRenderWorker(): Promise<FakeSelf> {
  const fake = installSelf();
  vi.resetModules();
  await import('./render-worker.js');
  return fake;
}

beforeEach(() => {
  initMock.mockReset();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('pptx render-worker.ts — init failure never hangs a request (AR4)', () => {
  it('a parse after a REJECTED init responds with an error (not a hang)', async () => {
    initMock.mockRejectedValue(new Error('render wasm boom'));
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({ data: { kind: 'parse', id: 9, buffer: new ArrayBuffer(4) } } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'error')).toBe(true);
    });

    const err = fake.posted.find((m) => (m as { kind?: string }).kind === 'error') as {
      id: number;
      message: string;
    };
    expect(err.id).toBe(9);
    expect(err.message).toContain('boom');
  });

  it('a parse after a SUCCESSFUL init responds with parsedMeta and no ready handshake', async () => {
    initMock.mockResolvedValue(undefined);
    const fake = await loadRenderWorker();

    fake.onmessage?.({ data: { kind: 'init', wasmUrl: 'x' } } as MessageEvent);
    fake.onmessage?.({ data: { kind: 'parse', id: 2, buffer: new ArrayBuffer(4) } } as MessageEvent);
    await vi.waitFor(() => {
      expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'parsedMeta')).toBe(true);
    });

    expect(fake.posted.some((m) => (m as { kind?: string }).kind === 'ready')).toBe(false);
  });
});
