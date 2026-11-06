// One harness lane (`lanes.ts`): a harness page, and one `LaneJob` at a time from stdin. Each
// result goes to stdout as a `LaneMessage` line. The lane exits when stdin ends, or with code 1
// when no new browser opens after a crash (the result of that deck is sent first).

import { rmSync } from 'node:fs';
import { withHarnessPage } from './harness-page.ts';
import { LANE_MESSAGE, type LaneJob, type LaneMessage } from './lanes.ts';
import { scoreSafely } from './score-deck.ts';

/**
 * The lines of stdin. Not `node:readline`: with it under Bun, the lane exited with code 0 in the
 * middle of `browser.close()` after stdin ended (no `beforeExit`, so an explicit exit whose
 * caller is unknown), and left the browser and the harness server running.
 */
async function* stdinLines(): AsyncGenerator<string> {
  const decoder = new TextDecoder();
  let buffered = '';
  for await (const chunk of Bun.stdin.stream()) {
    buffered += decoder.decode(chunk, { stream: true });
    let end: number;
    while ((end = buffered.indexOf('\n')) >= 0) {
      yield buffered.slice(0, end);
      buffered = buffered.slice(end + 1);
    }
  }
  if (buffered) yield buffered;
}

const send = (message: LaneMessage): void => {
  process.stdout.write(`${LANE_MESSAGE}${JSON.stringify(message)}\n`);
};

try {
  await withHarnessPage(async (first, control) => {
    let page = first;
    let started = false;
    send({ chromium: page.context().browser()?.version() ?? null });
    for await (const line of stdinLines()) {
      const job = JSON.parse(line) as LaneJob;
      if (job.imagesDir) rmSync(job.imagesDir, { recursive: true, force: true });
      // A corpus deck gets the fonts that it names, so each one opens in a new browser context:
      // its scores do not depend on the fonts of the decks before it. Every fixture deck gets
      // the same fonts (`localFonts(undefined, 'fixtures')`), so the fixture decks of a lane
      // share one page, which loads each font once (`harness.ts`) and scores as a new page does.
      if (started && job.input.fontSet !== 'fixtures') page = await control.restart();
      started = true;
      const scored = await scoreSafely(page, control, job.input, {
        imagesDir: job.imagesDir,
        redact: job.redact,
      });
      page = scored.page;
      send({ result: scored.result, fatal: scored.fatal?.message });
      if (scored.fatal) throw scored.fatal;
    }
  });
} catch (e) {
  console.error(`fidelity: a harness lane stopped: ${(e as Error).message.split('\n')[0]}`);
  process.exit(1);
}
process.exit(0);
