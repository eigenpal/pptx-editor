// The waits of `usePptxSource`: a URL that stops sending bytes fails, a slow download that keeps
// sending bytes finishes, and eager fonts that never settle hold the deck only for the wait.
import { afterEach, describe, expect, test } from 'bun:test';
import { fetchPptxSource, isAbortError, watchEagerPptxFonts } from '../pptx-source.ts';

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
const STALL = 60;
type Fonts = { supportedFamilies: string[] };

/** Answer with a body that `write` streams. The fake ignores the abort, as a bad host can. */
function serve(write: (c: ReadableStreamDefaultController<Uint8Array>) => Promise<void>): void {
  globalThis.fetch = (async () =>
    new Response(
      new ReadableStream<Uint8Array>({
        // A cancelled read closes the stream: later writes fail, as a server write would.
        start: (c) => void write(c).catch(() => undefined),
      })
    )) as unknown as typeof fetch;
}

async function settle(run: Promise<Uint8Array>): Promise<{ ms: number; result: unknown }> {
  const t0 = performance.now();
  const result = await run.then(
    (bytes) => bytes,
    (cause: unknown) => cause
  );
  return { ms: performance.now() - t0, result };
}

const fetchWithStall = (signal = new AbortController().signal) =>
  settle(fetchPptxSource('https://example.test/deck.pptx', undefined, signal, STALL));

describe('fetchPptxSource', () => {
  test('a request that never answers fails with a timeout, not an abort', async () => {
    globalThis.fetch = (() => new Promise<Response>(() => {})) as unknown as typeof fetch;
    const { ms, result } = await fetchWithStall();
    expect((result as Error).name).toBe('TimeoutError');
    expect(isAbortError(result)).toBe(false);
    expect(ms).toBeLessThan(STALL * 10);
  });

  test('a body that stops halfway fails with a timeout', async () => {
    serve(async (c) => c.enqueue(new Uint8Array(10)));
    const { ms, result } = await fetchWithStall();
    expect((result as Error).name).toBe('TimeoutError');
    expect(ms).toBeLessThan(STALL * 10);
  });

  test('a slow body that keeps sending bytes finishes past the stall limit', async () => {
    serve(async (c) => {
      for (let i = 0; i < 8; i += 1) {
        c.enqueue(new Uint8Array([i]));
        await sleep(STALL / 3);
      }
      c.close();
    });
    const { ms, result } = await fetchWithStall();
    expect(result).toEqual(new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7]));
    expect(ms).toBeGreaterThan(STALL);
  });

  test("the caller's abort ends a stalled body as an abort", async () => {
    serve(async (c) => c.enqueue(new Uint8Array(10)));
    const controller = new AbortController();
    setTimeout(() => controller.abort(), 10);
    const { ms, result } = await fetchWithStall(controller.signal);
    expect(isAbortError(result)).toBe(true);
    expect(ms).toBeLessThan(STALL);
  });

  test('an HTTP error fails at once', async () => {
    globalThis.fetch = (async () => new Response('', { status: 404 })) as unknown as typeof fetch;
    const { result } = await fetchWithStall();
    expect(String(result)).toContain('404');
  });
});

describe('watchEagerPptxFonts', () => {
  test('an origin that never settles releases the deck after the wait', async () => {
    const calls: string[] = [];
    const warn = console.warn;
    console.warn = () => undefined;
    try {
      watchEagerPptxFonts(
        () => new Promise<Fonts>(() => {}),
        { resolved: () => void calls.push('resolved'), release: () => void calls.push('release') },
        STALL
      );
      await sleep(STALL / 2);
      expect(calls).toEqual([]);
      await sleep(STALL);
      expect(calls).toEqual(['release']);
    } finally {
      console.warn = warn;
    }
  });

  test('fonts that settle after the release still apply, and release only once', async () => {
    const calls: string[] = [];
    let answer: (fonts: Fonts) => void = () => undefined;
    const warn = console.warn;
    console.warn = () => undefined;
    try {
      watchEagerPptxFonts(
        () => new Promise<Fonts>((resolve) => (answer = resolve)),
        {
          resolved: (fonts) => void calls.push(`resolved ${fonts.supportedFamilies?.join()}`),
          release: () => void calls.push('release'),
        },
        STALL
      );
      await sleep(STALL * 2);
      answer({ supportedFamilies: ['Acme Sans'] });
      await sleep(10);
      expect(calls).toEqual(['release', 'resolved Acme Sans']);
    } finally {
      console.warn = warn;
    }
  });

  test('fonts that settle in time apply before the release', async () => {
    const calls: string[] = [];
    watchEagerPptxFonts(
      async () => ({ supportedFamilies: ['Acme Sans'] }),
      { resolved: () => void calls.push('resolved'), release: () => void calls.push('release') },
      STALL
    );
    await sleep(10);
    expect(calls).toEqual(['resolved', 'release']);
  });

  test('a cancel stops both', async () => {
    const calls: string[] = [];
    const cancel = watchEagerPptxFonts(
      async () => ({ supportedFamilies: ['Acme Sans'] }),
      { resolved: () => void calls.push('resolved'), release: () => void calls.push('release') },
      STALL
    );
    cancel();
    await sleep(STALL * 2);
    expect(calls).toEqual([]);
  });
});
