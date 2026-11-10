import { expect, test } from 'bun:test';
import { waitForPublishedPackages } from './registry-ready.mjs';
import { names } from './common.mjs';
const metadata = {
  version: '0.4.0',
  gitHead: 'commit',
  dist: { integrity: 'sha512-value', tarball: 'https://registry.npmjs.org/package.tgz' },
};
const ready = () => Response.json(metadata);
const options = { pause: async () => {}, log: () => {} };

test('waits for every package and does not refetch ready packages', async () => {
  const calls = new Map();
  let pauses = 0;
  await waitForPublishedPackages('0.4.0', {
    ...options,
    pause: async () => {
      pauses++;
    },
    request: async (url) => {
      const name = decodeURIComponent(new URL(url).pathname.split('/')[1]);
      calls.set(name, (calls.get(name) ?? 0) + 1);
      return name === names[1] && calls.get(name) === 1
        ? new Response('', { status: 404 })
        : ready();
    },
  });
  expect(pauses).toBe(1);
  expect([...calls.values()]).toEqual([1, 2, 1]);
});

for (const status of [404, 429, 503, 'network']) {
  test(`temporary ${status} failure stops at the retry budget`, async () => {
    let pauses = 0;
    let requests = 0;
    await expect(
      waitForPublishedPackages('0.4.0', {
        ...options,
        attempts: 2,
        pause: async () => {
          pauses++;
        },
        request: async () => {
          requests++;
          if (status === 'network') throw new Error('offline');
          return new Response('', { status });
        },
      })
    ).rejects.toThrow('after 2 attempts');
    expect(pauses).toBe(1);
    expect(requests).toBe(names.length * 2);
  });
}

test('authentication failures and invalid metadata fail without retry', async () => {
  const pause = async () => {
    throw new Error('must not wait');
  };
  await expect(
    waitForPublishedPackages('0.4.0', {
      pause,
      request: async () => new Response('', { status: 403 }),
    })
  ).rejects.toThrow('HTTP 403');
  await expect(
    waitForPublishedPackages('0.4.0', {
      pause,
      request: async () => Response.json({ version: '0.4.0' }),
    })
  ).rejects.toThrow('Incomplete registry metadata');
});
