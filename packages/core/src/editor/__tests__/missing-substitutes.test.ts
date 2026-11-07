// Two loads of the substitutes that later slides need, at once (the opening lane and late
// shaping): the second waits for the first, so each face loads once.

import { expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { openPresentation } from '../document-session.ts';
import { continueOpening } from '../progressive-open.ts';

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

async function openedWhole(deck: Uint8Array) {
  const r = await openPresentation(deck, { progressive: { startSlide: 0 } });
  if (!r.ok) throw new Error('The deck did not open');
  const opening = r.session.opening;
  expect(opening).not.toBeNull();
  await new Promise<void>((done) =>
    continueOpening(opening!, {
      current: () => true,
      publish: () => {},
      settled: () => done(),
      failed: () => done(),
    })
  );
  return r.session;
}

test('concurrent loads of missing substitutes add each face once', async () => {
  // A large deck opens progressively: the open picks substitutes from the first slide only.
  const deck = buildPptx({
    slides: [
      slide('a', 'Arial'),
      slide('b'.repeat(700_000), 'Times New Roman'),
      slide('c', 'Courier New'),
    ],
  });
  const once = await openedWhole(deck);
  expect(await once.addMissingSubstitutes()).toBe(true);
  const twice = await openedWhole(deck);
  const results = await Promise.all([twice.addMissingSubstitutes(), twice.addMissingSubstitutes()]);
  expect(results).toEqual([true, false]);
  expect(twice.fontReport.faces.length).toBe(once.fontReport.faces.length);
  once.dispose();
  twice.dispose();
});
