// Packaged substitutes that do not arrive in time: the deck opens with the fixed measurer
// after the font wait, and measures again with shaped text when they land, as DOCX Editor
// opens before its fonts.
//
// The module mock needs its own process: `bun run test` gives each file one.

import { expect, mock, test } from 'bun:test';
import * as realFonts from '@docx-editor.dev/fonts';

const original = realFonts.loadDefaultFonts;
const copy = { ...realFonts };
// The substitutes load when the test calls `release`.
let release: () => void = () => {};
let held = Promise.resolve();
const hold = (): void => void (held = new Promise<void>((resolve) => (release = resolve)));
mock.module('@docx-editor.dev/fonts', () => ({
  ...copy,
  loadDefaultFonts: (o: Parameters<typeof realFonts.loadDefaultFonts>[0] = {}) =>
    held.then(() => original(o)),
}));

const settle = async (editor: { fontMeasurement(): { resolving: boolean } }) => {
  for (let i = 0; i < 400 && editor.fontMeasurement().resolving; i += 1)
    await new Promise((r) => setTimeout(r, 5));
};

test('stalled substitutes open the deck with the fixed measurer, then shaped text', async () => {
  const { createPptxEditor } = await import('../index.ts');
  const { FONT_WAIT_MS } = await import('../host-fonts.ts');
  const { buildPptx, para, run, shape } =
    await import('../../../../../scripts/fixtures/builder.ts');
  const text = para(run('Hi'));
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400, text }) }],
  });
  hold();
  const editor = createPptxEditor();
  const start = performance.now();
  expect((await editor.load(deck)).ok).toBe(true);
  expect(performance.now() - start).toBeLessThan(FONT_WAIT_MS + 1000);
  expect(editor.snapshot().status).toBe('ready');
  expect(editor.fontMeasurement()).toEqual({ measurer: 'fixed', resolving: true });
  const generation = editor.session!.fontGeneration;
  release();
  await settle(editor);
  expect(editor.fontMeasurement()).toEqual({ measurer: 'shaped', resolving: false });
  expect(editor.session!.fontGeneration).toBeGreaterThan(generation);
  editor.destroy();
}, 10_000);
