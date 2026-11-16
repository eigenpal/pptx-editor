// An editor open waits only for the packaged substitute faces that the text of the slide it
// shows uses. The other faces and the substitutes of the other slides load after the open and
// install when they arrive, so a deck does not hold its first slide for every font file.

import { afterEach, expect, test } from 'bun:test';
import {
  deckNaming,
  hold,
  release,
  requestedFaces,
  resetGate,
  waitFor,
} from './substitute-gate.ts';

const { createPptxEditor } = await import('../index.ts');
const { openPresentation } = await import('../document-session.ts');

afterEach(resetGate);

test('the open waits only for the substitutes of the slide it shows', async () => {
  hold('Times New Roman');
  hold('Courier New');
  const editor = createPptxEditor();
  expect(
    (await editor.load(await deckNaming([undefined, 'Times New Roman', 'Courier New']))).ok
  ).toBe(true);
  const session = editor.session!;
  // The first slide measures with shaped text at once.
  expect(editor.fontMeasurement()).toEqual({ measurer: 'shaped', resolving: true });
  // The theme names Calibri, and the master's bullet font (`a:buFont`) names Arial.
  expect([...session.substituteFamilies!].sort()).toEqual(['Arial', 'Calibri']);
  // The families still on the way are not reported as substitutions yet.
  expect(editor.snapshot().substitutedFonts).toEqual([]);
  const generation = session.fontGeneration;

  release('Times New Roman');
  release('Courier New');
  await waitFor(() => !editor.fontMeasurement().resolving);
  expect([...session.substituteFamilies!].sort()).toEqual([
    'Arial',
    'Calibri',
    'Courier New',
    'Times New Roman',
  ]);
  expect(session.fontGeneration).toBeGreaterThan(generation);
  expect(editor.snapshot().substitutedFonts.map((f) => f.family)).toContain('Times New Roman');
  // Each face loads once: the regular faces that the opening slide's text uses first, then the
  // other faces of those families, and the families of the other slides.
  expect(requestedFaces.slice(0, 2).sort()).toEqual(['Arial 400 normal', 'Calibri 400 normal']);
  const rest = requestedFaces.slice(2);
  expect(rest).toHaveLength(14);
  expect(new Set(rest).size).toBe(14);
  expect(rest).toContain('Calibri 700 normal');
  expect(rest).toContain('Courier New 700 italic');
  editor.destroy();
});

test('an open on another slide waits for the substitutes of that slide', async () => {
  hold('Cambria');
  const editor = createPptxEditor();
  const deck = await deckNaming([undefined, 'Cambria', 'Century Gothic']);
  expect((await editor.load(deck, { startSlide: 2 })).ok).toBe(true);
  const session = editor.session!;
  expect(session.substituteFamilies).toContain('Century Gothic');
  expect(session.substituteFamilies).not.toContain('Cambria');
  expect(editor.fontMeasurement().resolving).toBe(true);
  release('Cambria');
  await waitFor(() => !editor.fontMeasurement().resolving);
  expect(session.substituteFamilies).toContain('Cambria');
  editor.destroy();
});

test('an open that waits for its fonts loads every substitute before it opens', async () => {
  // A PDF export or a server session opens without `shapingWaitMs`: the open does not split.
  const deck = await deckNaming([undefined, 'Times New Roman', 'Courier New']);
  const opened = await openPresentation(deck, { firstSlideFonts: 0 });
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  expect([...session.substituteFamilies!].sort()).toEqual([
    'Arial',
    'Calibri',
    'Courier New',
    'Times New Roman',
  ]);
  expect(session.restSubstitutes).toBeNull();
  session.dispose();
});

test('a deck whose slides use only the opening slide families loads their other faces later', async () => {
  const opened = await openPresentation(await deckNaming([undefined, undefined]), {
    shapingWaitMs: 5000,
    firstSlideFonts: 0,
  });
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  expect(await session.restSubstitutes).toBe(true);
  // The families stay the same; only the bold and italic faces were added.
  expect([...session.substituteFamilies!].sort()).toEqual(['Arial', 'Calibri']);
  expect(session.measurer.resolve({ family: 'Calibri', bold: true, italic: false })).toMatchObject({
    family: 'Carlito',
    bold: true,
  });
  session.dispose();
});
