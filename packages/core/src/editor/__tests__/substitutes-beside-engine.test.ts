// The packaged substitutes download next to the text engine, not after it. On a slow network
// the two waits would otherwise add up, and the deck would stay on the fixed measurer longer.
//
// The module mock needs its own process: `bun run test` gives each file one.

import { expect, mock, test } from 'bun:test';
import * as realFonts from '@docx-editor.dev/fonts';

const original = realFonts.loadDefaultFonts;
const copy = { ...realFonts };
let requested = 0;
mock.module('@docx-editor.dev/fonts', () => ({
  ...copy,
  loadDefaultFonts: (o: Parameters<typeof realFonts.loadDefaultFonts>[0] = {}) => {
    requested += 1;
    return original(o);
  },
}));

test('the substitutes start downloading before the text engine is ready', async () => {
  const { openPresentation } = await import('../document-session.ts');
  const { buildPptx, para, run, shape } =
    await import('../../../../../scripts/fixtures/builder.ts');
  const text = para(run('Hi'));
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400, text }) }],
  });
  // No wait: the open returns at the end of the parse, before the text engine has loaded.
  const opened = await openPresentation(deck, { shapingWaitMs: 0 });
  if (!opened.ok) throw new Error('open failed');
  expect(opened.session.shaping).not.toBeNull();
  expect(requested).toBeGreaterThan(0);
  expect(await opened.session.shaping).toBe(true);
  opened.session.dispose();
});
