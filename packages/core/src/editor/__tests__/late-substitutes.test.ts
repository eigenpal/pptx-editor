// Packaged substitutes that arrive after a progressive open: they load for every slide that
// parsed while the measurer was fixed, not only for the slides the open had parsed.
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

test('late substitutes also load for slides that parsed while the measurer was fixed', async () => {
  const { createPptxEditor } = await import('../index.ts');
  const { buildPptx, para, run, shape } =
    await import('../../../../../scripts/fixtures/builder.ts');
  const slide = (text: string, face: string) => ({
    shapes: shape({
      id: 2,
      x: 0,
      y: 0,
      cx: 914400,
      cy: 914400,
      text: para(run(text, '', `<a:latin typeface="${face}"/>`)),
    }),
  });
  // A large deck opens progressively: the open picks substitutes from the first slide only.
  const deck = buildPptx({
    slides: [
      slide('a', 'Arial'),
      slide('b'.repeat(700_000), 'Times New Roman'),
      slide('c', 'Courier New'),
    ],
  });
  hold();
  const editor = createPptxEditor();
  expect((await editor.load(deck)).ok).toBe(true);
  const session = editor.session!;
  for (let i = 0; i < 400 && session.opening; i += 1) await new Promise((r) => setTimeout(r, 5));
  expect(session.opening).toBeNull();
  release();
  await settle(editor);
  expect(editor.fontMeasurement()).toEqual({ measurer: 'shaped', resolving: false });
  // Every family the parsed slides name has its substitute now: nothing is left to load.
  expect(await session.addMissingSubstitutes()).toBe(false);
  editor.destroy();
}, 15_000);
