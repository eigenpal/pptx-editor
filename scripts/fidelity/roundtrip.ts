// Round-trip mode: open → save → reopen every deck of a job list with the editor's own engine,
// once as is and once after a scripted edit (`roundtrip-census.ts`). The interface that
// `eigenpal/pptx-eval` calls for comparison R (`CONTRIBUTING.md`, "Fidelity harness").
//
// It needs no browser. Output is JSON only and holds no paths, file names, or
// deck text. The caller owns the policy.

import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import {
  appendFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { join, resolve } from 'node:path';
import { parseArgs } from 'node:util';
import {
  openResults,
  parseJobList,
  parseShard,
  redactPaths,
  shardJobs,
  sourcesHash,
  writeJsonAtomic,
  type CorpusJob,
} from './corpus-list.ts';
import { sha256 } from './reference-identity.ts';
import {
  censusDeck,
  EDIT_MOVE_EMU,
  EDIT_RECIPE,
  EDIT_TEXT,
  exception,
  MachineFault,
  ROUNDTRIP_SCHEMA,
  type CensusStage,
  type DeckCensus,
  type PassOneFacts,
} from './roundtrip-census.ts';

export const ROUNDTRIP_USAGE = `Usage:
  bun scripts/fidelity/roundtrip.ts --list <jobs.json|jobs.jsonl> --out <dir> [--saved <dir>] [--shard i/n] [--resume | --fresh]
      Round-trip mode for eigenpal/pptx-eval comparison R (CONTRIBUTING.md, "Fidelity harness").
      --list <file>    the job list of corpus mode: {id, deck, pdf} per line, deck an absolute
                       path; pdf is ignored
      --out <dir>      results.jsonl, run.json, and current.json; give each shard its own folder
      --saved <dir>    keep each saved copy as <id>.saved.pptx and <id>.edited.pptx
      --shard i/n      run every n-th job, starting at job i
      --resume         skip finished decks; refuses a folder of another engine, uncommitted
                       changes, job list, or shard, a change between keeping copies (--saved)
                       and not, and a folder with results but no readable run.json
      --fresh          replace the results of an earlier run in --out and start a new run;
                       without --resume or --fresh, a non-empty results.jsonl stops the run
      -h, --help       print this text

Exit codes: 0 the run finished, 2 a usage error, 3 a fault of the machine (a folder, a kept
copy, run.json, current.json, or results.jsonl could not be written; the deck in progress runs
again on --resume), 1 a crash of the tool (on --resume, the deck in progress is recorded as
interrupted).`;

/** Write one of the run's own files. A failure is a fault of the machine, never of a deck. */
function write(what: string, action: () => void): void {
  try {
    action();
  } catch (e) {
    throw new MachineFault(`cannot write ${what} (${(e as { code?: string }).code ?? 'error'})`);
  }
}

/** A usage error: one line on stderr and exit code 2. */
class UsageError extends Error {}
/** Exit code for a fault of the machine (a full disk): the run stops, and the deck in progress
 * runs again on --resume once the machine is fixed. */
export const MACHINE_FAULT_EXIT = 3;

/**
 * The paths whose source the round trip runs: the engine packages, these scripts, and the
 * dependency pins. An edit elsewhere (docs, fixtures) does not change the engine.
 */
export const ENGINE_PATHS = [
  'packages/core',
  'packages/ooxml-core',
  'packages/i18n',
  'scripts/fidelity',
  'package.json',
  'bun.lock',
  // The census reads the OOXML schemas when it runs: they set the element names in the output.
  'reference/ecma-376/part4/transitional-schemas',
];

/**
 * The files that decide a round-trip result, hashed from disk as corpus mode's `sourcesHash`
 * does. This hash decides a resume, and it works without git; the git fields below name the
 * revision for a reader. Tests are left out.
 */
export const ROUNDTRIP_SOURCES = [
  'packages/core/src',
  'packages/ooxml-core/src',
  'packages/i18n/src',
  'scripts/fidelity',
  'reference/ecma-376/part4/transitional-schemas',
  'package.json',
  'tsconfig.json',
  'bun.lock',
] as const;
const REPO_ROOT = resolve(import.meta.dir, '../..');

/**
 * The commit of this checkout and a hash of the uncommitted changes in the engine paths,
 * tracked and untracked, wherever the command runs from. Null when git is not available.
 */
export function gitState(cwd: string = import.meta.dir): {
  commit: string | null;
  dirty: boolean | null;
  changes: string | null;
} {
  const git = (...args: string[]): string =>
    execFileSync('git', args, {
      cwd,
      encoding: 'utf8',
      maxBuffer: 256 * 1024 ** 2,
    });
  try {
    const commit = git('rev-parse', 'HEAD').trim();
    const root = git('rev-parse', '--show-toplevel').trim();
    const paths = ENGINE_PATHS.map((path) => `:(top)${path}`);
    const diff = git('diff', 'HEAD', '--no-ext-diff', '--binary', '--', ...paths);
    const untracked = git(
      'ls-files',
      '--others',
      '--exclude-standard',
      '--full-name',
      '-z',
      '--',
      ...paths
    )
      .split('\0')
      .filter(Boolean)
      .sort();
    if (diff === '' && untracked.length === 0) return { commit, dirty: false, changes: null };
    const hash = createHash('sha256').update(diff);
    for (const path of untracked) {
      hash.update(`\0${path}\0`);
      // An unreadable file (a broken link, a nested repository) still counts, by its error.
      try {
        hash.update(readFileSync(join(root, path)));
      } catch (e) {
        hash.update(`unreadable:${(e as { code?: string }).code ?? 'error'}`);
      }
    }
    return { commit, dirty: true, changes: hash.digest('hex') };
  } catch {
    return { commit: null, dirty: null, changes: null };
  }
}

/** One progress line: the error, or each check of both passes. */
export function progressLine(r: DeckCensus): string {
  // A digest or compare error keeps the facts of the steps before it: show both.
  const error = r.error ? `${r.error.stage} ${r.error.reason}` : '';
  if (r.error && !r.roundTrip) return error;
  const t = r.roundTrip;
  const e = r.edit;
  const pass1 = t
    ? `differences ${t.differences.count}, namespace ${t.namespaceDifferences.count}, digest ${t.digestStable === null ? '-' : t.digestStable ? 'stable' : 'unstable'}, bytes ${t.binaryParts}/${t.xmlParts + t.binaryParts}`
    : 'pass 1 -';
  if (!e) return `${error ? `${error}; ` : ''}${pass1}; edit -`;
  const edits = e.edits.map((x) => `${x.edit} ${x.result}`).join(', ') || 'none';
  const undo = e.undoRestores === null ? '-' : e.undoRestores ? 'ok' : 'no';
  const kept =
    e.fingerprintEqual === null ? '-' : e.fingerprintEqual ? 'yes' : `no ${e.differences.count}`;
  return `${error ? `${error}; ` : ''}${pass1}; ${edits}, kept ${kept}, collateral ${e.collateral.count}, lost ${e.lost}, undo ${undo}${e.error ? `, ${e.error.stage} ${e.error.reason}` : ''}`;
}

const empty: DeckCensus = { error: null, slides: null, parts: null, roundTrip: null, edit: null };

const EMPTY_EDIT = {
  edits: [],
  error: null,
  editedParts: 0,
  fingerprintEqual: null,
  differences: { count: 0, items: [] },
  collateral: { count: 0, items: [] },
  lost: 0,
  undoRestores: null,
};

interface Options {
  readonly listText: string;
  readonly jobs: CorpusJob[];
  readonly shard: { index: number; count: number };
  readonly out: string;
  readonly saved: string | undefined;
  readonly resume: boolean;
  readonly fresh: boolean;
}

/** Parse and check every argument before the command writes anything. */
function options(args: readonly string[]): Options | 'help' {
  let parsed;
  try {
    parsed = parseArgs({
      args: [...args],
      strict: true,
      allowPositionals: false,
      options: {
        list: { type: 'string' },
        out: { type: 'string' },
        saved: { type: 'string' },
        shard: { type: 'string' },
        resume: { type: 'boolean' },
        fresh: { type: 'boolean' },
        help: { type: 'boolean', short: 'h' },
      },
    }).values;
  } catch (e) {
    throw new UsageError(e instanceof Error ? e.message : String(e));
  }
  if (parsed.help) return 'help';
  if (!parsed.list || !parsed.out) throw new UsageError('--list and --out are required');
  if (parsed.resume && parsed.fresh)
    throw new UsageError('--resume and --fresh exclude each other');
  for (const [flag, value] of [
    ['--out', parsed.out],
    ['--saved', parsed.saved],
    ['--shard', parsed.shard],
  ] as const) {
    if (value === undefined) continue;
    if (value === '' || value.startsWith('-')) throw new UsageError(`${flag} needs a value`);
    if (flag !== '--shard' && existsSync(value) && !statSync(value).isDirectory())
      throw new UsageError(`${flag} must be a folder, not a file`);
  }
  let listText: string;
  try {
    listText = readFileSync(parsed.list, 'utf8');
  } catch (e) {
    throw new UsageError(`cannot read the job list (${(e as { code?: string }).code ?? 'error'})`);
  }
  try {
    const shard = parsed.shard ? parseShard(parsed.shard) : { index: 1, count: 1 };
    return {
      listText,
      jobs: shardJobs(parseJobList(listText), shard),
      shard,
      out: resolve(parsed.out),
      saved: parsed.saved ? resolve(parsed.saved) : undefined,
      resume: parsed.resume === true,
      fresh: parsed.fresh === true,
    };
  } catch (e) {
    throw new UsageError(e instanceof Error ? e.message : String(e));
  }
}

function readJson(path: string): Record<string, unknown> | undefined {
  try {
    return JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
  } catch {
    return undefined;
  }
}

async function main(args: readonly string[]): Promise<void> {
  const parsed = options(args);
  if (parsed === 'help') {
    console.log(ROUNDTRIP_USAGE);
    return;
  }
  const { listText, jobs: listed, shard, out, saved: savedDir, resume, fresh } = parsed;
  const git = gitState();
  const identity = {
    schema: ROUNDTRIP_SCHEMA,
    ...git,
    sources: sourcesHash(REPO_ROOT, ROUNDTRIP_SOURCES),
    list: sha256(new TextEncoder().encode(listText)),
    shard,
    keepsCopies: savedDir !== undefined,
  };
  const resultsFile = join(out, 'results.jsonl');
  const runFile = join(out, 'run.json');
  // The deck in progress and its step. A deck that crashes or hangs the process is recorded on
  // resume, once, instead of running again.
  const currentFile = join(out, 'current.json');
  const before = readJson(runFile);
  // A folder with results but no readable run.json (a write cut by a full disk, or a removed
  // file) cannot prove its identity, so it is refused too.
  if (resume && (before || existsSync(currentFile) || existsSync(resultsFile))) {
    // A resumed run must come from the same engine, uncommitted changes, job list, and shard, or
    // one results file would mix two engines.
    if (before === undefined)
      throw new UsageError(
        'the output folder has no readable run.json, so its run cannot be checked; use a new --out'
      );
    const names: Record<string, string> = {
      schema: 'schema',
      commit: 'commit',
      dirty: 'uncommitted changes',
      changes: 'uncommitted changes',
      sources: 'engine sources',
      list: 'job list',
      shard: '--shard',
      keepsCopies: '--saved',
    };
    const differ = [
      ...new Set(
        (Object.keys(identity) as (keyof typeof identity)[])
          .filter((key) => JSON.stringify(before[key]) !== JSON.stringify(identity[key]))
          .map((key) => names[key])
      ),
    ];
    if (differ.length)
      throw new UsageError(
        `these differ from the run in the output folder: ${differ.join(', ')}; use a new --out`
      );
  }
  // As corpus mode: a long run's results are never replaced by a missing flag.
  if (!resume && !fresh && existsSync(resultsFile) && statSync(resultsFile).size > 0)
    throw new UsageError(
      'results.jsonl exists in --out. Add --resume to continue that run, or --fresh to replace it'
    );
  write('the --out folder', () => mkdirSync(out, { recursive: true }));
  if (savedDir) write('the --saved folder', () => mkdirSync(savedDir, { recursive: true }));
  let done = new Set<string>();
  write('results.jsonl', () => {
    done = openResults(resultsFile, resume);
  });
  const ids = new Set(listed.map((job) => job.id));
  // The deck that an earlier run stopped on, recorded below as `interrupted`.
  let interrupted: CensusStage | null = null;
  const current = readJson(currentFile) as
    { id?: string; stage?: CensusStage; facts?: PassOneFacts } | undefined;
  if (resume && current?.id && ids.has(current.id) && !done.has(current.id)) {
    const error = { stage: current.stage ?? 'open', reason: 'interrupted' };
    interrupted = error.stage;
    // A deck that stopped in the edit step keeps the facts of its finished pass 1.
    const line =
      current.stage === 'edit' && current.facts
        ? { ...empty, ...current.facts, edit: { ...EMPTY_EDIT, error } }
        : { ...empty, error };
    write('results.jsonl', () =>
      appendFileSync(
        resultsFile,
        `${JSON.stringify({ id: current.id, milliseconds: null, ...line })}\n`
      )
    );
    // A copy that the interrupted deck left may be partial.
    if (savedDir)
      for (const kind of ['saved', 'edited'])
        write('the --saved folder', () =>
          rmSync(join(savedDir, `${current.id}.${kind}.pptx`), { force: true })
        );
    write('results.jsonl', () => {
      done = openResults(resultsFile, true);
    });
  }
  write('current.json', () => rmSync(currentFile, { force: true }));
  const jobs = listed.filter((j) => !done.has(j.id));
  const runInfo = {
    ...identity,
    tool: 'pptx-editor scripts/fidelity/roundtrip.ts',
    bun: Bun.version,
    edit: { recipe: EDIT_RECIPE, setText: EDIT_TEXT, moveEmu: EDIT_MOVE_EMU },
    // A resumed run keeps its start; a fresh run (also one that replaces results) starts now.
    startedAt:
      (resume ? (before?.startedAt as string | undefined) : undefined) ?? new Date().toISOString(),
    finishedAt: null as string | null,
    decks: { total: listed.length, skipped: done.size, finished: 0, errors: 0, editErrors: 0 },
  };
  if (interrupted !== null && current?.id) {
    // The interrupted deck is a result of this invocation, not a skipped one.
    runInfo.decks.skipped -= 1;
    runInfo.decks.finished += 1;
    if (interrupted === 'edit') runInfo.decks.editErrors += 1;
    else runInfo.decks.errors += 1;
    console.log(`${current.id}: ${interrupted} interrupted (an earlier run stopped on this deck)`);
  }
  const writeRun = (): void => write('run.json', () => writeJsonAtomic(runFile, runInfo));
  writeRun();
  for (const job of jobs) {
    let stage: CensusStage = 'open';
    const enter = (next: CensusStage, facts?: PassOneFacts): void => {
      stage = next;
      write('current.json', () =>
        writeFileSync(currentFile, JSON.stringify({ id: job.id, stage, facts }))
      );
    };
    const copy = (kind: 'saved' | 'edited'): string => join(savedDir!, `${job.id}.${kind}.pptx`);
    // After a fault of the machine the deck runs again on --resume: remove its progress file and
    // any partial copies.
    const abandon = (): void => {
      // A failed removal (a read-only disk) must not hide the fault of the machine.
      try {
        rmSync(currentFile, { force: true });
        if (savedDir)
          for (const kind of ['saved', 'edited'] as const) rmSync(copy(kind), { force: true });
      } catch {
        // The fault in progress is reported instead.
      }
    };
    enter('open');
    if (savedDir)
      for (const kind of ['saved', 'edited'] as const)
        write('the --saved folder', () => rmSync(copy(kind), { force: true }));
    const keep = savedDir
      ? (kind: 'saved' | 'edited', bytes: Uint8Array): void => {
          try {
            writeFileSync(copy(kind), bytes);
          } catch (e) {
            rmSync(copy(kind), { force: true });
            throw new MachineFault(
              `cannot write a kept copy (${(e as { code?: string }).code ?? 'error'})`
            );
          }
        }
      : undefined;
    const started = performance.now();
    let r: DeckCensus;
    try {
      r = await censusDeck(new Uint8Array(readFileSync(job.deck)), { keep, onStage: enter });
    } catch (e) {
      if (e instanceof MachineFault) {
        abandon();
        throw e;
      }
      // A deck that cannot be read, or a failure outside the census itself, is an error of this
      // deck; the run goes on.
      const code = (e as { code?: unknown }).code;
      r =
        stage === 'open' && (code === 'ENOENT' || code === 'EISDIR' || code === 'EACCES')
          ? { ...empty, error: { stage, reason: code === 'ENOENT' ? 'missing' : 'unreadable' } }
          : { ...empty, error: exception(stage, e) };
    }
    const milliseconds = Math.round(performance.now() - started);
    try {
      write('results.jsonl', () =>
        appendFileSync(resultsFile, `${JSON.stringify({ id: job.id, milliseconds, ...r })}\n`)
      );
    } catch (e) {
      // A cut line is removed on --resume, and the deck runs again.
      abandon();
      throw e;
    }
    write('current.json', () => rmSync(currentFile, { force: true }));
    runInfo.decks.finished += 1;
    if (r.error) runInfo.decks.errors += 1;
    if (r.edit?.error) runInfo.decks.editErrors += 1;
    if (runInfo.decks.finished % 25 === 0) writeRun();
    console.log(`${job.id}: ${progressLine(r)}`);
  }
  runInfo.finishedAt = new Date().toISOString();
  writeRun();
  const { finished, errors, editErrors, skipped } = runInfo.decks;
  console.log(
    `Decks: ${finished} run, ${errors} with a pass-1 error, ${editErrors} with a pass-2 error, ${skipped} skipped. Results: ${resultsFile}`
  );
}

if (import.meta.main) {
  try {
    await main(process.argv.slice(2));
  } catch (e) {
    if (e instanceof UsageError) {
      console.error(
        `roundtrip: ${redactPaths(e.message)} (bun scripts/fidelity/roundtrip.ts --help)`
      );
      process.exit(2);
    }
    if (e instanceof MachineFault) {
      console.error(
        `roundtrip: ${e.message}; the run stops. Fix the machine, then run again with --resume`
      );
      process.exit(MACHINE_FAULT_EXIT);
    }
    throw e;
  }
}
