// A slide picture waits for the page's fonts only as long as the face load wait: a face whose
// load never settles keeps the browser's font set loading, and must not hold an export.
import { afterEach, expect, test } from 'bun:test';
import { FACE_LOAD_WAIT_MS } from '../../output/font-faces.ts';
import { fontsSettled } from '../export-image.ts';

const g = globalThis as { document?: unknown };
const realDocument = g.document;
afterEach(() => {
  g.document = realDocument;
});

test('a font set that never gets ready holds a slide picture for the face load wait only', async () => {
  g.document = { fonts: { ready: new Promise(() => {}) } };
  const t0 = performance.now();
  await fontsSettled();
  const ms = performance.now() - t0;
  expect(ms).toBeGreaterThanOrEqual(FACE_LOAD_WAIT_MS - 50);
  expect(ms).toBeLessThan(FACE_LOAD_WAIT_MS * 3);
});

test('a ready font set lets a slide picture go at once', async () => {
  g.document = { fonts: { ready: Promise.resolve() } };
  const t0 = performance.now();
  await fontsSettled();
  expect(performance.now() - t0).toBeLessThan(FACE_LOAD_WAIT_MS / 2);
});

test('with no document, a slide picture does not wait for fonts', async () => {
  g.document = undefined;
  await expect(fontsSettled()).resolves.toBeUndefined();
});

test('slides exported one after another wait for a stalled font set once, not once each', async () => {
  g.document = { fonts: { ready: new Promise(() => {}) } };
  await fontsSettled();
  const t0 = performance.now();
  for (let i = 0; i < 20; i += 1) await fontsSettled();
  expect(performance.now() - t0).toBeLessThan(FACE_LOAD_WAIT_MS / 2);
});
