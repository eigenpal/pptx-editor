// An open that stops waiting for its fonts measures with the stand-in measurer it was given
// (the browser canvas in an editor), not the fixed measurer, until shaping starts.

import { afterEach, expect, test } from 'bun:test';
import { createFixedMeasurer } from '../../layout/index.ts';
import { deckNaming, hold, release, resetGate } from './substitute-gate.ts';

const { openPresentation } = await import('../document-session.ts');

afterEach(resetGate);

test('a stalled open measures with the stand-in until shaping starts', async () => {
  hold('Calibri');
  hold('Arial');
  const standIn = createFixedMeasurer();
  const opened = await openPresentation(await deckNaming([undefined]), {
    shapingWaitMs: 0,
    standIn: () => standIn,
  });
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  expect(session.fixedMeasurer).toBe(true);
  expect(session.measurer).toBe(standIn);
  release('Calibri');
  release('Arial');
  expect(await session.shaping).toBe(true);
  expect(session.fixedMeasurer).toBe(false);
  expect(session.measurer).not.toBe(standIn);
  session.dispose();
});

test('a stand-in that cannot measure leaves the fixed measurer', async () => {
  // Calibri and Arial loaded in the test before: the opening slide names Cambria too.
  hold('Cambria');
  const opened = await openPresentation(await deckNaming(['Cambria']), {
    shapingWaitMs: 0,
    standIn: () => null,
  });
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  expect(session.fixedMeasurer).toBe(true);
  expect(session.measurer.resolve({ family: 'Calibri', bold: false, italic: false }).id).toBe(
    'Calibri'
  );
  release('Cambria');
  await session.shaping;
  session.dispose();
});
