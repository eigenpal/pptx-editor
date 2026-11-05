// The command line of the fidelity harness (`run.ts`). Flags are strict: an unknown flag, a
// missing value, or a flag of the other mode is a usage error (exit 2), never ignored.

import { parseArgs } from 'node:util';

export const RUN_USAGE = `Usage:
  bun run fidelity [-- --deck <name>] [--print-gaps]
      The fixture gate: every fixture in fixtures/ against its reference PDF and the known gaps.
      --deck <name>   one fixture only, for example tables
      --print-gaps    print the known-gap entries that this run's scores record
  bun run fidelity:corpus [-- --deck deckN]
      The private corpus in demo-presentations/ (local only): a report, no known gaps.
  bun scripts/fidelity/run.ts --list <jobs.json|jobs.jsonl> --out <dir> [--shard i/n] [--resume | --fresh] [--images]
      Corpus mode for eigenpal/pptx-eval (CONTRIBUTING.md, "Fidelity harness"). --resume continues
      the run in --out; --fresh replaces its results. Without either, existing results stop the run.

Exit codes: 0 pass, 1 the gate or a slide fails (corpus mode: the run could not finish), 2 a
usage error, 3 not gated (no known-gap list for this machine, outside CI).`;

export type RunArgs =
  | { readonly mode: 'help' }
  | {
      readonly mode: 'set';
      readonly set: 'fixtures' | 'corpus';
      readonly deck?: string;
      readonly printGaps: boolean;
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
  const { values: v } = parseArgs({
    args: [...argv],
    strict: true,
    allowPositionals: false,
    options: {
      set: { type: 'string' },
      deck: { type: 'string' },
      'print-gaps': { type: 'boolean' },
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
    const other = (['set', 'deck', 'print-gaps'] as const).find((k) => v[k] !== undefined);
    if (other) throw new Error(`--${other} does not go with --list`);
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
  if (v['print-gaps'] && set !== 'fixtures')
    throw new Error('--print-gaps goes only with the fixtures');
  return { mode: 'set', set, deck: v.deck, printGaps: v['print-gaps'] ?? false };
}
