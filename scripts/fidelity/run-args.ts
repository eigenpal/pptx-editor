// The command line of the fidelity harness (`run.ts`). Flags are strict: an unknown flag, a
// missing value, or a flag of the other mode is a usage error (exit 2), never ignored.

import { parseArgs } from 'node:util';

export const RUN_USAGE = `Usage:
  bun run fidelity [<deck>...] [--update-gaps [--accept-worse]]
      Print each fixture in fixtures/ with the editor, and compare it with its reference PDF.
      <deck>          only these fixtures, for example: bun run fidelity tables
      --update-gaps   write this run's failing slides to FIXTURE_KNOWN_GAPS in thresholds.ts
      --accept-worse  also write new entries and looser values (a new fixture, a changed deck,
                      a new reference PDF, or new tools)
  bun run fidelity:corpus [<deckN>...]
      The private corpus in demo-presentations/ (local only): a report, no known gaps.
  bun scripts/fidelity/run.ts --list <jobs.json|jobs.jsonl> --out <dir> [--shard i/n] [--resume | --fresh] [--images]
      Corpus mode for eigenpal/pptx-eval. --resume continues the run in --out; --fresh replaces
      its results. Without either, existing results stop the run.

Exit codes: 0 pass, 1 the gate or a slide fails (corpus mode: the run could not finish), 2 a
usage error.`;

export type RunArgs =
  | { readonly mode: 'help' }
  | {
      readonly mode: 'set';
      readonly set: 'fixtures' | 'corpus';
      /** Only these decks; empty for all. */
      readonly decks: readonly string[];
      readonly updateGaps: boolean;
      /** With `updateGaps`: write new entries and looser values too. */
      readonly acceptWorse: boolean;
    }
  | {
      readonly mode: 'list';
      readonly list: string;
      readonly out: string;
      readonly shard?: string;
      readonly resume: boolean;
      readonly fresh: boolean;
      readonly images: boolean;
    };

/** Parse the harness's arguments. Throws an Error with the usage problem; the caller exits 2. */
export function parseRunArgs(argv: readonly string[]): RunArgs {
  const { values: v, positionals } = parseArgs({
    args: [...argv],
    strict: true,
    allowPositionals: true,
    options: {
      set: { type: 'string' },
      'update-gaps': { type: 'boolean' },
      'accept-worse': { type: 'boolean' },
      list: { type: 'string' },
      out: { type: 'string' },
      shard: { type: 'string' },
      resume: { type: 'boolean' },
      fresh: { type: 'boolean' },
      images: { type: 'boolean' },
      help: { type: 'boolean', short: 'h' },
    },
  });
  if (v.help) return { mode: 'help' };
  if (v.list !== undefined) {
    if (v.set !== undefined) throw new Error('--set does not go with --list');
    if (v['update-gaps']) throw new Error('--update-gaps does not go with --list');
    if (v['accept-worse']) throw new Error('--accept-worse does not go with --list');
    if (positionals.length) throw new Error('deck names do not go with --list');
    if (!v.out) throw new Error('--list needs --out <dir>');
    if (v.resume && v.fresh) throw new Error('--resume and --fresh do not go together');
    return {
      mode: 'list',
      list: v.list,
      out: v.out,
      shard: v.shard,
      resume: v.resume ?? false,
      fresh: v.fresh ?? false,
      images: v.images ?? false,
    };
  }
  const listOnly = (['out', 'shard', 'resume', 'fresh', 'images'] as const).find(
    (k) => v[k] !== undefined
  );
  if (listOnly) throw new Error(`--${listOnly} goes only with --list`);
  const set = v.set ?? 'fixtures';
  if (set !== 'fixtures' && set !== 'corpus')
    throw new Error(`--set must be fixtures or corpus, not ${set}`);
  if (v['update-gaps'] && set !== 'fixtures')
    throw new Error('--update-gaps goes only with the fixtures');
  if (v['accept-worse'] && !v['update-gaps'])
    throw new Error('--accept-worse goes only with --update-gaps');
  return {
    mode: 'set',
    set,
    decks: positionals,
    updateGaps: v['update-gaps'] ?? false,
    acceptWorse: v['accept-worse'] ?? false,
  };
}
