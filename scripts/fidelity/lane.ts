// One harness lane (`lanes.ts`): a harness page, and one `LaneJob` at a time from stdin. Each
// result goes to stdout as a `LaneMessage` line. The lane exits when stdin ends, or with code 1
// when no new browser opens after a crash (the result of that deck is sent first).

import { rmSync } from 'node:fs';
import { createInterface } from 'node:readline';
import { withHarnessPage } from './harness-page.ts';
import { LANE_MESSAGE, type LaneJob, type LaneMessage } from './lanes.ts';
import { scoreSafely } from './score-deck.ts';

const send = (message: LaneMessage): void => {
  process.stdout.write(`${LANE_MESSAGE}${JSON.stringify(message)}\n`);
};

try {
  await withHarnessPage(async (first, control) => {
    let page = first;
    let started = false;
    send({ chromium: page.context().browser()?.version() ?? null });
    for await (const line of createInterface({ input: process.stdin })) {
      const job = JSON.parse(line) as LaneJob;
      if (job.imagesDir) rmSync(job.imagesDir, { recursive: true, force: true });
      // Each deck in a new browser context, so its scores do not depend on the decks before it.
      if (started) page = await control.restart();
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
