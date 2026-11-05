// Dev probe: render one slide of one deck to PNG through the harness page.
//   bun scripts/fidelity/render-probe.ts <deck.pptx> <slideIndex> <out.png>
import { writeFileSync } from 'node:fs';
import { openDeck, renderSlide, withHarnessPage } from './harness-page.ts';

const [deck, index, out] = process.argv.slice(2);
await withHarnessPage(async (page) => {
  page.on('console', (m) => console.log('[page]', m.text()));
  const { slides, error } = await openDeck(page, deck);
  console.log({ slides, error });
  const r = await renderSlide(page, Number(index), 1920);
  console.log(JSON.stringify(r.diagnostics));
  writeFileSync(out, r.png);
});
