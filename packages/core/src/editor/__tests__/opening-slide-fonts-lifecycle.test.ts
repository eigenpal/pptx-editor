// The substitutes that an editor open leaves for later: host origins skip them, and they do
// nothing once the editor that asked for them is gone.

import { afterEach, expect, test } from 'bun:test';
import { deckNaming, hold, release, resetGate, waitFor } from './substitute-gate.ts';

const { createPptxEditor, defineFontResolver } = await import('../index.ts');

afterEach(resetGate);

test('substitutes that arrive after the editor is destroyed do nothing', async () => {
  hold('Times New Roman');
  const editor = createPptxEditor();
  expect((await editor.load(await deckNaming([undefined, 'Times New Roman']))).ok).toBe(true);
  const session = editor.session!;
  editor.destroy();
  release('Times New Roman');
  expect(await session.restSubstitutes).toBe(false);
  expect(session.substituteFamilies).not.toContain('Times New Roman');
});

test('a host origin does not load the substitutes that load after the open', async () => {
  hold('Courier New');
  const asked: string[] = [];
  const editor = createPptxEditor({
    fonts: defineFontResolver((request) => {
      asked.push(...(request.resolvedFaces ?? []).map((f) => f.family));
      return undefined;
    }),
  });
  expect((await editor.load(await deckNaming([undefined, 'Courier New']))).ok).toBe(true);
  // The engine loads Liberation Mono for Courier New after the open: an origin skips it.
  expect(asked).toContain('Liberation Mono');
  release('Courier New');
  await waitFor(() => !editor.fontMeasurement().resolving);
  editor.destroy();
});

test('a later deck replaces one whose substitutes are still on the way', async () => {
  hold('Cambria');
  const editor = createPptxEditor();
  expect((await editor.load(await deckNaming([undefined, 'Cambria']))).ok).toBe(true);
  const first = editor.session!;
  expect((await editor.load(await deckNaming([undefined]))).ok).toBe(true);
  const generation = editor.session!.fontGeneration;
  release('Cambria');
  await first.restSubstitutes;
  await waitFor(() => !editor.fontMeasurement().resolving);
  // The replaced deck's late faces do not touch the deck on screen.
  expect(editor.session!.fontGeneration).toBe(generation);
  editor.destroy();
});
