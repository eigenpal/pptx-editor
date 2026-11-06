// Score decks in parallel lanes. Each lane is a child process (`lane.ts`) with its own harness
// server, Chromium, and CPU for the comparison, and scores one deck at a time with
// `scoreSafely`, as a sequential run does. The lanes take decks from one queue, and the results
// come back in the order of the input, so the report does not depend on which lane was faster.

import { cpus } from 'node:os';
import { join } from 'node:path';
import type { DeckResult } from './report.ts';
import type { DeckInput } from './score-deck.ts';

/** One deck for a lane. */
export interface LaneJob {
  readonly input: DeckInput;
  /** Where the deck's images go (`scoreSafely`); the lane clears it first. */
  readonly imagesDir?: string;
  readonly redact: boolean;
}

/** A message from a lane: the first gives the Chromium version, each later one a deck. */
export type LaneMessage =
  { readonly chromium: string | null } | { readonly result: DeckResult; readonly fatal?: string };

/** The start of each stdout line of a lane that carries a `LaneMessage`. */
export const LANE_MESSAGE = '\u0000fidelity-lane ';

/**
 * The number of lanes: `PPTX_FIDELITY_JOBS` when set, else one for each core, at most 4. A lane
 * waits on its browser for much of a deck, so it does not fill a core. On a 14-core laptop the
 * fixtures took 83 s with 2 lanes, 35 s with 4, and 38 s with 6 or 10, before the save check
 * compared markup and the raster cache kept the filtered pages. Never more lanes than
 * decks.
 */
// Not `availableParallelism`: under Bun it reads `navigator`, which the pdf.js legacy build
// replaces, and then it gives undefined.
export function laneCount(decks: number, env = process.env.PPTX_FIDELITY_JOBS): number {
  const asked = env ? Number(env) : NaN;
  const lanes =
    Number.isInteger(asked) && asked > 0 ? asked : Math.min(4, Math.max(1, cpus().length));
  return Math.max(1, Math.min(lanes, decks));
}

/**
 * Score `jobs` in `lanes` child processes. `onResult` runs as each deck finishes. Resolves with
 * the results in the order of `jobs` and the Chromium version. Rejects when a lane could not
 * score at all (it did not start, or no new browser opened after a crash): every finished deck
 * was passed to `onResult` first.
 */
export async function scoreInLanes(
  jobs: readonly LaneJob[],
  lanes: number,
  onResult: (job: LaneJob, result: DeckResult) => void
): Promise<{ results: DeckResult[]; chromium: string | null }> {
  const results: DeckResult[] = new Array(jobs.length);
  let next = 0;
  let chromium: string | null = null;
  const children = new Set<Bun.Subprocess<'pipe', 'pipe', 'inherit'>>();
  let failed: Error | undefined;

  const lane = async (): Promise<void> => {
    const child = Bun.spawn([process.execPath, join(import.meta.dir, 'lane.ts')], {
      stdin: 'pipe',
      stdout: 'pipe',
      stderr: 'inherit',
    });
    children.add(child);
    let current: number | undefined;
    const give = (): void => {
      if (failed || next >= jobs.length) {
        current = undefined;
        void child.stdin.end();
        return;
      }
      current = next;
      next += 1;
      child.stdin.write(`${JSON.stringify(jobs[current])}\n`);
      void child.stdin.flush();
    };
    const handle = (line: string): void => {
      if (!line.startsWith(LANE_MESSAGE)) {
        if (line) console.log(line);
        return;
      }
      const message = JSON.parse(line.slice(LANE_MESSAGE.length)) as LaneMessage;
      if ('chromium' in message) {
        chromium ??= message.chromium;
        give();
        return;
      }
      if (current === undefined) return;
      results[current] = message.result;
      onResult(jobs[current], message.result);
      if (message.fatal) {
        failed ??= new Error(message.fatal);
        current = undefined;
      } else give();
    };
    let buffered = '';
    const decoder = new TextDecoder();
    const reader = child.stdout.getReader();
    for (let chunk = await reader.read(); !chunk.done; chunk = await reader.read()) {
      buffered += decoder.decode(chunk.value, { stream: true });
      let end: number;
      while ((end = buffered.indexOf('\n')) >= 0) {
        handle(buffered.slice(0, end));
        buffered = buffered.slice(end + 1);
      }
    }
    if (buffered) handle(buffered);
    const code = await child.exited;
    children.delete(child);
    if (failed) throw failed;
    if (current !== undefined || code !== 0)
      throw new Error(`a harness lane exited (code ${code}) before it finished its decks`);
  };

  try {
    await Promise.all(Array.from({ length: lanes }, lane));
  } catch (e) {
    failed ??= e as Error;
    for (const c of children) c.kill(9);
    throw failed;
  }
  return { results, chromium };
}
