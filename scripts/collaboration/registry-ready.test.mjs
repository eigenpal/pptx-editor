import { expect, test } from 'bun:test';
import { waitForPublishedPackages, PUBLICATION_TIMEOUT_MS } from './registry-ready.mjs';
import { names } from './common.mjs';
const metadata = {
  version: '0.4.0',
  gitHead: 'commit',
  dist: { integrity: 'sha512-value', tarball: 'https://registry.npmjs.org/package.tgz' },
};
const ready = () => Response.json(metadata);
function clock() {
  let time = 0;
  const pauses = [];
  return {
    pauses,
    now: () => time,
    random: () => 1,
    log: () => {},
    pause: async (ms) => {
      pauses.push(ms);
      time += ms;
    },
  };
}

test('waits for every exact version, including propagation beyond five minutes', async () => {
  const time = clock();
  const calls = new Map();
  await waitForPublishedPackages('0.4.0', {
    ...time,
    request: async (url, options) => {
      expect(url.endsWith('/0.4.0')).toBe(true);
      expect(options.cache).toBe('no-store');
      const name = decodeURIComponent(new URL(url).pathname.split('/')[1]);
      calls.set(name, (calls.get(name) ?? 0) + 1);
      return name === names[1] && time.now() < 360_000
        ? new Response('', { status: 404 })
        : ready();
    },
  });
  expect(calls.get(names[0])).toBe(1);
  expect(calls.get(names[2])).toBe(1);
  expect(time.now()).toBeGreaterThanOrEqual(360_000);
  expect(time.now()).toBeLessThan(PUBLICATION_TIMEOUT_MS);
});

for (const status of [404, 408, 429, 500, 502, 503, 504, 'network', 'body', 'cancel']) {
  test(`temporary ${status} failure stops at one shared deadline`, async () => {
    const time = clock();
    await expect(
      waitForPublishedPackages('0.4.0', {
        ...time,
        timeoutMs: 12_000,
        request: async () => {
          if (status === 'network') throw new Error('offline');
          if (status === 'cancel')
            return {
              ok: false,
              status: 503,
              headers: new Headers(),
              body: {
                cancel: async () => {
                  throw new Error('disconnected');
                },
              },
            };
          if (status === 'body')
            return {
              ok: true,
              json: async () => {
                throw new Error('body disconnected');
              },
            };
          return new Response('', { status });
        },
      })
    ).rejects.toThrow('before deadline');
    expect(time.now()).toBe(12_000);
    expect(time.pauses).toEqual([5000, 7000]);
  });
}

for (const value of ['120', new Date(120_000).toUTCString()]) {
  test(`honors Retry-After ${value}`, async () => {
    const time = clock();
    await waitForPublishedPackages('0.4.0', {
      ...time,
      request: async () =>
        time.now() === 0
          ? new Response('', { status: 429, headers: { 'retry-after': value } })
          : ready(),
    });
    expect(time.pauses).toEqual([120_000]);
  });
}

test('Retry-After cannot extend the shared deadline', async () => {
  const time = clock();
  await expect(
    waitForPublishedPackages('0.4.0', {
      ...time,
      timeoutMs: 1000,
      request: async () => new Response('', { status: 503, headers: { 'retry-after': '120' } }),
    })
  ).rejects.toThrow('before deadline');
  expect(time.now()).toBe(1000);
});

for (const status of [400, 401, 403, 501]) {
  test(`permanent HTTP ${status} fails immediately`, async () => {
    const time = clock();
    await expect(
      waitForPublishedPackages('0.4.0', {
        ...time,
        request: async () => new Response('', { status }),
      })
    ).rejects.toThrow(`HTTP ${status}`);
    expect(time.pauses).toEqual([]);
  });
}

test('invalid metadata cancels other pending requests without retry', async () => {
  const time = clock();
  const signals = [];
  await expect(
    waitForPublishedPackages('0.4.0', {
      ...time,
      request: async (_url, { signal }) => {
        signals.push(signal);
        if (signals.length === 1) return Response.json({ version: '99.0.0' });
        return new Promise((_resolve, reject) =>
          signal.addEventListener('abort', () => reject(signal.reason), { once: true })
        );
      },
    })
  ).rejects.toThrow('Incomplete registry metadata');
  expect(signals.every((signal) => signal.aborted)).toBe(true);
  expect(time.pauses).toEqual([]);
});
