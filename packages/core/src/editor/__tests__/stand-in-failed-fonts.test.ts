// When every font fails to load (a blocked network, a strict proxy), the stand-in measurer
// keeps measuring after shaping starts, so the text never falls back to half-em estimates.

import { afterEach, expect, test } from 'bun:test';
import { createFixedMeasurer } from '../../layout/index.ts';
import { deckNaming, fail, hold, release, resetGate } from './substitute-gate.ts';

const { openPresentation } = await import('../document-session.ts');

afterEach(resetGate);

test('when every font fails to load, the stand-in keeps measuring after shaping starts', async () => {
  hold('Calibri');
  fail('Calibri');
  fail('Arial');
  const standIn = createFixedMeasurer();
  const opened = await openPresentation(await deckNaming([undefined]), {
    shapingWaitMs: 0,
    standIn: () => standIn,
  });
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  release('Calibri');
  await session.shaping;
  // No face arrived: the browser's faces measure, not half-em estimates.
  expect(session.measurer).toBe(standIn);
  session.dispose();
});
