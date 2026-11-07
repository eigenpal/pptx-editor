// The demo's deck download: a stall fails soon, and a slow download that keeps sending bytes
// finishes, however long it takes in all.
import { afterEach, describe, expect, test } from 'bun:test';
import { DeckStallError, fetchDeck } from './demo-files.ts';

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
const STALL = 60;

/** Answer with a body that `write` streams. The fake ignores the abort, as a bad host can. */
function serve(write: (c: ReadableStreamDefaultController<Uint8Array>) => Promise<void>): void {
  globalThis.fetch = (async () =>
    new Response(new ReadableStream<Uint8Array>({ start: (c) => void write(c) }), {
      status: 200,
    })) as unknown as typeof fetch;
}

async function settle(run: Promise<Uint8Array>): Promise<{ ms: number; result: unknown }> {
  const t0 = performance.now();
  const result = await run.then(
    (bytes) => bytes,
    (cause: unknown) => cause
  );
  return { ms: performance.now() - t0, result };
}

describe('fetchDeck', () => {
  test('a request that never answers fails after the stall limit', async () => {
    globalThis.fetch = (() => new Promise<Response>(() => {})) as unknown as typeof fetch;
    const { ms, result } = await settle(fetchDeck('/', 'deck.pptx', STALL));
    expect(result).toBeInstanceOf(DeckStallError);
    expect(ms).toBeLessThan(STALL * 10);
  });

  test('a body that stops halfway fails after the stall limit', async () => {
    serve(async (c) => c.enqueue(new Uint8Array(10)));
    const { ms, result } = await settle(fetchDeck('/', 'deck.pptx', STALL));
    expect(result).toBeInstanceOf(DeckStallError);
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
    const { ms, result } = await settle(fetchDeck('/', 'deck.pptx', STALL));
    expect(result).toEqual(new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7]));
    expect(ms).toBeGreaterThan(STALL);
  });

  test('an HTTP error fails at once', async () => {
    globalThis.fetch = (async () => new Response('', { status: 404 })) as unknown as typeof fetch;
    const { result } = await settle(fetchDeck('/', 'deck.pptx', STALL));
    expect(String(result)).toContain('HTTP 404');
  });
});
