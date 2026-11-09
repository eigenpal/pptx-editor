// The publish that reports a deck as ready already counts late shaped measurement in
// `resolving`. A host that waits for the fonts before it lifts its loading screen reads
// `fontMeasurement()` in that publish; it must not see the fonts as settled while the packaged
// substitutes are still on the way.
//
// The module mock needs its own process: `bun run test` gives each file one.

import { expect, mock, test } from 'bun:test';
import * as realFonts from '@docx-editor.dev/fonts';

const original = realFonts.loadDefaultFonts;
const copy = { ...realFonts };
let release: () => void = () => {};
const held = new Promise<void>((resolve) => (release = resolve));
mock.module('@docx-editor.dev/fonts', () => ({
  ...copy,
  loadDefaultFonts: (o: Parameters<typeof realFonts.loadDefaultFonts>[0] = {}) =>
    held.then(() => original(o)),
}));

test('the ready publish reports late fonts as resolving', async () => {
  const { createPptxEditor } = await import('../index.ts');
  const { buildPptx, para, run, shape } =
    await import('../../../../../scripts/fixtures/builder.ts');
  const text = para(run('Hi'));
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400, text }) }],
  });
  const editor = createPptxEditor();
  // What the first publish with the deck ready reported.
  const seen: { atReady?: ReturnType<typeof editor.fontMeasurement> } = {};
  const off = editor.subscribe(() => {
    if (!seen.atReady && editor.snapshot().status === 'ready' && editor.session)
      seen.atReady = editor.fontMeasurement();
  });
  expect((await editor.load(deck)).ok).toBe(true);
  expect(seen.atReady).toEqual({ measurer: 'fixed', resolving: true });
  release();
  for (let i = 0; i < 400 && editor.fontMeasurement().resolving; i += 1)
    await new Promise((r) => setTimeout(r, 5));
  expect(editor.fontMeasurement()).toEqual({ measurer: 'shaped', resolving: false });
  off();
  editor.destroy();
}, 10_000);
