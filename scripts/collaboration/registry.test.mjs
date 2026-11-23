import { expect, test } from 'bun:test';
import { registry, REGISTRY_RETRY_DELAYS_MS } from './common.mjs';

const notFound = () => {
  throw new Error('npm failed: npm error code E404');
};

// Post-release updates run 45246799486: `npm view` got a 404 one second after the registry
// wait passed. A short retry covers a registry edge that still serves the older packument.
test('retries a 404 from npm view, then returns the metadata', () => {
  const pauses = [];
  let calls = 0;
  const metadata = registry('@pptx-editor.dev/core', '0.7.0', {
    view: (spec) => {
      expect(spec).toBe('@pptx-editor.dev/core@0.7.0');
      if (++calls < 3) notFound();
      return JSON.stringify({ version: '0.7.0' });
    },
    pause: (ms) => pauses.push(ms),
  });
  expect(metadata).toEqual({ version: '0.7.0' });
  expect(pauses).toEqual(REGISTRY_RETRY_DELAYS_MS.slice(0, 2));
});

test('stops after the last retry', () => {
  const pauses = [];
  expect(() =>
    registry('@pptx-editor.dev/core', '0.7.0', { view: notFound, pause: (ms) => pauses.push(ms) })
  ).toThrow('E404');
  expect(pauses).toEqual(REGISTRY_RETRY_DELAYS_MS);
});

test('does not retry other failures', () => {
  const pauses = [];
  expect(() =>
    registry('@pptx-editor.dev/core', '0.7.0', {
      view: () => {
        throw new Error('npm failed: npm error code E401');
      },
      pause: (ms) => pauses.push(ms),
    })
  ).toThrow('E401');
  expect(pauses).toEqual([]);
});
