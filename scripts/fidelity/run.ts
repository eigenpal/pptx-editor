// The fidelity harness.
//
//   bun run fidelity          # synthetic fixtures (CI gate)
//   bun run fidelity:corpus   # the private corpus in demo-presentations/ (local only)
//   bun scripts/fidelity/run.ts --list jobs.jsonl --out <dir> [--shard i/n] [--resume] [--images]
//                             # corpus mode for eigenpal/pptx-eval (CONTRIBUTING.md, "Fidelity harness")
//   bun scripts/fidelity/run.ts --help
//
// Renders every slide with the editor's own pipeline in Playwright Chromium and compares it with
// the reference PDF (`score-deck.ts`).
//
// Fixture and corpus sets write an HTML report to fidelity-report/<set>/. The fixture set exits 1
// when the fixture gate fails (`known-gaps.ts`: a failure that is not a known gap, a known gap
// that got worse or is gone, a missing or stale reference, a deck the harness cannot score, or
// a save-consistency failure). The corpus set exits 1 when any slide with a current reference
// fails or is not save-consistent. Corpus mode writes JSON only and exits 0 when every deck was
// scored: the caller owns the policy. Every mode exits 2 on a usage error.

import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, resolve } from 'node:path';
import { FIXTURE_KNOWN_GAPS } from './thresholds.ts';
import {
  checkFixtureGate,
  fidelityEnvironment,
  gateExitCode,
  listToolChanges,
  recordedGaps,
  slideKey,
  slideVerdict,
} from './known-gaps.ts';
import { withHarnessPage } from './harness-page.ts';
import { laneCount, scoreInLanes, type LaneJob } from './lanes.ts';
import {
  corpusThresholds,
  CORPUS_SCHEMA,
  parseJobList,
  parseShard,
  redactPaths,
  resultLine,
  shardJobs,
  sourcesHash,
  startCorpusRun,
  writeJsonAtomic,
  type CorpusJob,
} from './corpus-list.ts';
import { localFontsIdentity, referenceFontsInstalled } from './local-fonts.ts';
import { popplerVersion, pruneRasterCache } from './references.ts';
import { deckFacts, readManifest, referenceState, sha256 } from './reference-identity.ts';
import { writeReport, type DeckResult } from './report.ts';
import { parseRunArgs, RUN_USAGE, type RunArgs } from './run-args.ts';
import { scoreSafely, type DeckInput } from './score-deck.ts';

const counts = (r: DeckResult): Record<string, number> =>
  r.slides.reduce(
    (m, s) => ({ ...m, [s.status]: (m[s.status] ?? 0) + 1 }),
    {} as Record<string, number>
  );

/**
 * The fixture rasters kept between runs: by PDF, Poppler version, and `pdftoppm` options
 * (`rasterCacheKey`). CI restores and saves this folder.
 */
const RASTER_CACHE = process.env.PPTX_FIDELITY_RASTER_CACHE || '.cache/fidelity-rasters';

const POPPLER_MISSING =
  'Poppler is not installed (pdftoppm, pdftotext, pdfinfo), so no slide can be compared. Install it: brew install poppler, or apt-get install poppler-utils.';

/** `bun run fidelity` and `bun run fidelity:corpus`: decks in a folder, an HTML report. */
async function runSet(args: Extract<RunArgs, { mode: 'set' }>): Promise<number> {
  const { set, deck: only } = args;
  const dir = set === 'corpus' ? 'demo-presentations' : 'fixtures';
  const out = resolve('fidelity-report', set);
  if (!existsSync(dir)) {
    console.log(`No decks in ${dir}/.`);
    return set === 'fixtures' ? 1 : 0;
  }
  const poppler = popplerVersion();
  if (!poppler && set === 'fixtures') {
    console.error(`GATE: ${POPPLER_MISSING}`);
    return 1;
  }
  const manifest = readManifest(dir);
  const files = readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.pptx'))
    .sort();
  const inputs: DeckInput[] = files.map((f, i) => {
    const file = join(dir, f);
    const name = set === 'corpus' ? `deck${i + 1}` : f.replace(/\.pptx$/, '');
    if (!poppler)
      return { file, name, gap: { status: 'no-reference', reasons: ['Poppler is not installed'] } };
    const state = referenceState(file, manifest);
    if (state.status === 'current') return { file, name, pdf: state.pdf };
    if (state.status === 'stale')
      return { file, name, gap: { status: 'stale', reasons: state.reasons } };
    return { file, name, gap: { status: 'no-reference', reasons: ['no reference PDF'] } };
  });
  const targets = only ? inputs.filter((d) => d.name === only) : inputs;
  if (only && targets.length === 0) {
    console.error(
      `fidelity: no ${set === 'corpus' ? 'corpus deck' : 'fixture'} is named ${only}. Names: ${inputs.map((d) => d.name).join(', ')}.`
    );
    return 2;
  }
  // The fixture rasters are kept between runs (`rasterizePages`); corpus rasters never are.
  const rasterCache = set === 'fixtures' ? resolve(RASTER_CACHE) : undefined;
  const started = Date.now();
  // The decks with the most slides start first, so no lane is left with a long deck at the end.
  const slideCount = (d: DeckInput): number => {
    try {
      return deckFacts(new Uint8Array(readFileSync(d.file))).slides;
    } catch {
      return 0;
    }
  };
  const jobs: LaneJob[] = targets
    .map((d) => ({ d, n: slideCount(d) }))
    .sort((a, b) => b.n - a.n)
    .map(({ d }) => ({
      input: rasterCache ? { ...d, rasterCache } : d,
      imagesDir: join(out, d.name),
      redact: false,
    }));
  const lanes = laneCount(jobs.length);
  console.log(`Scoring ${jobs.length} decks in ${lanes} lane(s).`);
  const scored = await scoreInLanes(jobs, lanes, (job, r) => {
    const consistent = r.slides.every((s) => s.selfConsistent);
    console.log(
      `${job.input.name}: ${r.error ?? JSON.stringify(counts(r))} save-consistent=${consistent}`
    );
  });
  // In the order of the decks, whatever lane finished first.
  const byName = new Map(scored.results.map((r, i) => [jobs[i].input.name, r]));
  const results = targets.map((d) => byName.get(d.name)!);
  const chromium = scored.chromium;
  if (rasterCache && !only) pruneRasterCache(rasterCache, started);
  if (set !== 'fixtures') {
    for (const r of results) for (const s of r.slides) s.verdict = slideVerdict(s, undefined);
    writeReport(out, set, results, { poppler, chromium });
    const failed = results.some(
      (r) => r.error || r.slides.some((s) => s.status === 'fail' || s.selfConsistent === false)
    );
    console.log(`Report: ${join(out, 'index.html')}`);
    return failed ? 1 : 0;
  }
  // The fixture gate: known gaps per environment, a ratchet.
  const referenceFonts = referenceFontsInstalled();
  const environment = fidelityEnvironment(process.platform, referenceFonts);
  const recorded = environment ? FIXTURE_KNOWN_GAPS[environment] : undefined;
  const gaps = recorded?.slides;
  const list = environment ? `FIXTURE_KNOWN_GAPS.${environment}` : undefined;
  const toolChanges = listToolChanges(recorded, poppler, chromium);
  const gate = checkFixtureGate(results, gaps, { complete: !only, list, toolChanges });
  const where = `${process.platform} ${referenceFonts ? 'with' : 'without'} reference fonts`;
  const problems = [...gate.problems];
  const notes = [...gate.notes];
  if (!environment && process.env.CI)
    problems.push(
      `no known-gap list for ${where} in CI. Add the environment to FIDELITY_ENVIRONMENTS and FIXTURE_KNOWN_GAPS in scripts/fidelity/thresholds.ts`
    );
  for (const r of results)
    for (const s of r.slides) {
      const key = slideKey(r.name, s.index);
      s.knownGap = gate.knownGaps.has(key);
      s.verdict = slideVerdict(s, gaps && Object.hasOwn(gaps, key) ? gaps[key] : undefined);
    }
  const entries = recordedGaps(results);
  writeReport(out, set, results, {
    poppler,
    chromium,
    gate: { where, list, problems, notes, entries },
  });
  const all = results.flatMap((r) => r.slides);
  console.log(
    `${all.filter((s) => s.status === 'pass').length} of ${all.length} slides pass; ${gate.knownGaps.size} fail as known gaps.`
  );
  if (list) console.log(`Known gaps: ${list}.`);
  for (const n of notes) console.log(`NOTE: ${n}`);
  for (const p of problems) console.log(`GATE: ${p}`);
  if (args.printGaps) {
    console.log(`The known-gap entries that this run's scores record (${where}):`);
    for (const e of entries) console.log(`    ${e}`);
  }
  console.log(`Report: ${join(out, 'index.html')}; each slide's scores are in summary.json.`);
  const code = gateExitCode(problems.length, Boolean(list));
  if (code === 1) console.log(`Gate: fails with ${problems.length} problem(s).`);
  else if (code === 3)
    // A distinct code, so a script cannot read a run that gated no slide score as a pass.
    console.log(
      `Gate: NOT GATED. There is no known-gap list for ${where}, so slide failures are reported, not gated; only references, open errors, and save consistency are gated. Run the gate in CI or on the Mac that exports the references (exit 3).`
    );
  else console.log(`Gate: passes${only ? ` (one fixture only: ${only})` : ''}.`);
  if (code === 1 && !list)
    console.log(
      `Slide scores were not gated here either: there is no known-gap list for ${where}.`
    );
  return code;
}

function gitState(): { commit: string | null; dirty: boolean | null } {
  try {
    const commit = execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim();
    const dirty =
      execFileSync('git', ['status', '--porcelain', '--untracked-files=no'], {
        encoding: 'utf8',
      }).trim() !== '';
    return { commit, dirty };
  } catch {
    return { commit: null, dirty: null };
  }
}

/** Corpus mode: a job list in, per-slide JSON out (`CONTRIBUTING.md`, "Fidelity harness"). */
async function runList(args: Extract<RunArgs, { mode: 'list' }>): Promise<number> {
  // Check every input before anything on disk changes.
  let listText: string;
  let jobs: CorpusJob[];
  let shard: { index: number; count: number };
  try {
    listText = readFileSync(args.list, 'utf8');
    shard = args.shard ? parseShard(args.shard) : { index: 1, count: 1 };
    jobs = parseJobList(listText);
  } catch (e) {
    console.error(`fidelity: ${redactPaths((e as Error).message)}\n\n${RUN_USAGE}`);
    return 2;
  }
  const poppler = popplerVersion();
  if (!poppler) {
    console.error(`fidelity: ${POPPLER_MISSING}`);
    return 1;
  }
  const out = resolve(args.out);
  mkdirSync(out, { recursive: true });
  const resultsFile = join(out, 'results.jsonl');
  const identity = {
    schema: CORPUS_SCHEMA,
    list: sha256(new TextEncoder().encode(listText)),
    shard,
    sources: sourcesHash(resolve('.')),
    thresholds: corpusThresholds(),
    popplerVersion: poppler,
    referenceFonts: referenceFontsInstalled(),
    fonts: localFontsIdentity(),
  };
  const start = startCorpusRun(out, identity, args);
  if ('refused' in start) {
    console.error(`fidelity: ${start.refused}`);
    return 2;
  }
  const mine = shardJobs(jobs, shard);
  const todo = mine.filter((j) => !start.done.has(j.id));
  const runInfo = {
    ...identity,
    harness: 'pptx-editor scripts/fidelity/run.ts',
    ...gitState(),
    poppler: true,
    /** PDF baselines are glyph origins from the content stream (pdf.js), else box estimates. */
    baselines: 'glyph-origin',
    chromium: null as string | null,
    startedAt: new Date().toISOString(),
    finishedAt: null as string | null,
    decks: { total: mine.length, skipped: mine.length - todo.length, scored: 0, errors: 0 },
  };
  // Through a rename, so a crash never leaves half a run.json for the next --resume.
  const writeRun = (): void => writeJsonAtomic(join(out, 'run.json'), runInfo);
  writeRun();
  await withHarnessPage(async (first, control) => {
    let page = first;
    runInfo.chromium = page.context().browser()?.version() ?? null;
    writeRun();
    for (const job of todo) {
      let r: DeckResult;
      let fatal: Error | undefined;
      if (!existsSync(job.deck)) r = { name: job.id, error: 'the deck does not exist', slides: [] };
      else {
        const imagesDir = args.images ? join(out, 'images', job.id) : undefined;
        if (imagesDir) rmSync(imagesDir, { recursive: true, force: true });
        const input: DeckInput =
          job.pdf && !existsSync(job.pdf)
            ? {
                file: job.deck,
                name: job.id,
                gap: { status: 'no-reference', reasons: ['the PDF does not exist'] },
              }
            : { file: job.deck, name: job.id, pdf: job.pdf ?? undefined };
        const scored = await scoreSafely(page, control, input, { imagesDir, redact: true });
        page = scored.page;
        r = scored.result;
        fatal = scored.fatal;
      }
      appendFileSync(resultsFile, `${resultLine(job, r)}\n`);
      runInfo.decks.scored += 1;
      if (r.error) runInfo.decks.errors += 1;
      writeRun();
      console.log(`${job.id}: ${r.error ?? JSON.stringify(counts(r))}`);
      if (fatal) throw fatal;
    }
  });
  runInfo.finishedAt = new Date().toISOString();
  writeRun();
  console.log(`Results: ${resultsFile}`);
  return 0;
}

async function main(argv: readonly string[]): Promise<number> {
  let args: RunArgs;
  try {
    args = parseRunArgs(argv);
  } catch (e) {
    console.error(`fidelity: ${(e as Error).message}\n\n${RUN_USAGE}`);
    return 2;
  }
  if (args.mode === 'help') {
    console.log(RUN_USAGE);
    return 0;
  }
  try {
    return await (args.mode === 'list' ? runList(args) : runSet(args));
  } catch (e) {
    // Not even a new browser opens the harness page, so no later deck can be scored. Every
    // finished deck is recorded.
    console.error(`fidelity: the harness stopped: ${(e as Error).message.split('\n')[0]}`);
    return 1;
  }
}

if (import.meta.main) process.exit(await main(process.argv.slice(2)));
