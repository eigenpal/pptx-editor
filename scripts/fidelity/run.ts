// The fidelity harness.
//
//   bun run fidelity                  # every fixture, both lanes: the gate
//   bun run fidelity tables           # one fixture
//   bun run fidelity --lane export    # one lane: editor (the editor's print) or export
//   bun run fidelity --update-gaps    # remove checks that pass now from thresholds.ts
//   bun run fidelity x --update-gaps --accept-worse   # also add new entries and kinds
//   bun run fidelity:corpus           # the private corpus in demo-presentations/ (local only)
//   bun scripts/fidelity/run.ts --list jobs.jsonl --out <dir> ...   # corpus mode for pptx-eval
//
// The editor prints each deck to PDF with its own paint, and `score-deck.ts` compares that PDF
// with the reference PDF. The export lane compares the PDF of `exportPdf` in the same way. Fixture
// and corpus sets write an HTML report, both PDFs, and the page images to fidelity-report/<set>/
// (the export lane: fidelity-report/<set>-export/). The fixture set exits 1 when the gate of a
// lane fails (`known-gaps.ts`); each lane has its own known-gap list.
// Corpus mode writes JSON only and exits 0 when every deck was scored: the caller owns the
// policy. Every mode exits 2 on a usage error.

import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, relative, resolve } from 'node:path';
import { FIXTURE_EXPORT_KNOWN_GAPS, FIXTURE_KNOWN_GAPS, type KnownGapList } from './thresholds.ts';
import {
  checkFixtureGate,
  planUpdate,
  slideKey,
  slideVerdict,
  writeGaps,
  type GapListName,
} from './known-gaps.ts';
import { pinnedChromium, withHarnessPage } from './harness-page.ts';
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
import {
  fixtureFontsIdentity,
  localFontsIdentity,
  referenceFontsInstalled,
} from './local-fonts.ts';
import { PDFJS_VERSION } from './pdf.ts';
import { deckFacts, readManifest, referenceState, sha256 } from './reference-identity.ts';
import { writeReport, type DeckResult } from './report.ts';
import { parseRunArgs, RUN_USAGE, type Lane, type RunArgs } from './run-args.ts';
import { scoreSafely, type DeckInput } from './score-deck.ts';

const THRESHOLDS_FILE = resolve(import.meta.dir, 'thresholds.ts');

const pct = (v: number): string => `${(v * 100).toFixed(2)} %`;

/**
 * One console line for a scored deck: slides that pass, failing slides that the known-gap list
 * names, hidden slides, the worst pixel score, and the time.
 */
function deckLine(r: DeckResult, list?: KnownGapList, prefix = ''): string {
  if (r.error) return `${prefix}${r.name}: not scored: ${r.error}`;
  const compared = r.slides.filter((s) => s.status === 'pass' || s.status === 'fail');
  const passing = compared.filter((s) => s.status === 'pass').length;
  const listed = list
    ? compared.filter(
        (s) => s.status === 'fail' && Object.hasOwn(list.slides, slideKey(r.name, s.index))
      ).length
    : 0;
  const known = listed ? `, ${listed} known ${listed === 1 ? 'gap' : 'gaps'}` : '';
  const hidden = r.slides.filter((s) => s.hidden).length;
  const hiddenText = hidden ? `, ${hidden} hidden` : '';
  const worst = Math.max(0, ...compared.map((s) => s.pixelMismatch ?? 0));
  const gap =
    r.reference === 'pdf' ? '' : ` (${r.reference === 'stale' ? 'stale' : 'no'} reference)`;
  const saved = r.slides.some((s) => s.selfConsistent === false) ? ', save/reopen DIFFERS' : '';
  const time = `${((r.milliseconds ?? 0) / 1000).toFixed(1)} s`;
  return `${prefix}${r.name}: ${passing}/${r.slides.length - hidden} pass${known}${hiddenText}, worst pixels ${pct(worst)}${gap}${saved}, ${time}`;
}

/** What differs between the two lanes of a run. */
const LANE_INFO: Record<
  Lane,
  {
    /** The report folder after `fidelity-report/<set>`. */
    readonly suffix: string;
    readonly gaps: GapListName;
    /** The start of each console line of the lane. */
    readonly prefix: string;
    /** The name of the lane in the report and the summary line. */
    readonly title: string;
    /** How our PDF is made, for the report. */
    readonly method: string;
  }
> = {
  editor: {
    suffix: '',
    gaps: 'FIXTURE_KNOWN_GAPS',
    prefix: '',
    title: 'editor print',
    method: "Our PDF is printed with the editor's own paint.",
  },
  export: {
    suffix: '-export',
    gaps: 'FIXTURE_EXPORT_KNOWN_GAPS',
    prefix: 'export ',
    title: 'PDF export',
    method:
      'Our PDF is the PDF of exportPdf (@pptx-editor.dev/pptx-to-pdf), with the fonts that the editor lane gets.',
  },
};

/** The known-gap list of a lane, as `thresholds.ts` holds it. */
const laneGaps = (lane: Lane): KnownGapList =>
  lane === 'export' ? FIXTURE_EXPORT_KNOWN_GAPS : FIXTURE_KNOWN_GAPS;

/** `bun run fidelity` and `bun run fidelity:corpus`: decks in a folder, an HTML report per lane. */
async function runSet(args: Extract<RunArgs, { mode: 'set' }>): Promise<number> {
  const { set, decks: only, lanes } = args;
  const dir = set === 'corpus' ? 'demo-presentations' : 'fixtures';
  const outOf = (lane: Lane): string =>
    resolve('fidelity-report', `${set}${LANE_INFO[lane].suffix}`);
  if (!existsSync(dir)) {
    console.log(`No decks in ${dir}/.`);
    return set === 'fixtures' ? 1 : 0;
  }
  const manifest = readManifest(dir);
  const files = readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.pptx'))
    .sort();
  const inputs: DeckInput[] = files.map((f, i) => {
    const file = join(dir, f);
    const name = set === 'corpus' ? `deck${i + 1}` : f.replace(/\.pptx$/, '');
    const state = referenceState(file, manifest);
    if (state.status === 'current') return { file, name, pdf: state.pdf };
    if (state.status === 'stale')
      return { file, name, gap: { status: 'stale', reasons: state.reasons } };
    return { file, name, gap: { status: 'no-reference', reasons: ['no reference PDF'] } };
  });
  const unknown = only.filter((n) => !inputs.some((d) => d.name === n));
  if (unknown.length) {
    console.error(
      `fidelity: no deck is named ${unknown.join(', ')}. Names: ${inputs.map((d) => d.name).join(', ')}.`
    );
    return 2;
  }
  const targets = only.length ? inputs.filter((d) => only.includes(d.name)) : inputs;
  if (!only.length) for (const lane of lanes) rmSync(outOf(lane), { recursive: true, force: true });
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
    .flatMap(({ d }) =>
      lanes.map((lane) => ({
        input: {
          ...d,
          ...(set === 'fixtures' ? { fontSet: 'fixtures' as const } : {}),
          ...(lane === 'export' ? { lane } : {}),
        },
        imagesDir: join(outOf(lane), d.name),
        redact: false,
      }))
    );
  const laneOf = (job: LaneJob): Lane => job.input.lane ?? 'editor';
  const started = Date.now();
  const scored = await scoreInLanes(jobs, laneCount(jobs.length), (job, r) => {
    const lane = laneOf(job);
    console.log(
      deckLine(r, set === 'fixtures' ? laneGaps(lane) : undefined, LANE_INFO[lane].prefix)
    );
  });
  const byKey = new Map(
    scored.results.map((r, i) => [`${laneOf(jobs[i])}\0${jobs[i].input.name}`, r])
  );
  const seconds = ((Date.now() - started) / 1000).toFixed(1);
  const tools = {
    pdfjs: PDFJS_VERSION,
    chromium: pinnedChromium(),
    fonts: set === 'fixtures' ? fixtureFontsIdentity() : localFontsIdentity(),
  };
  const toolsText = `pdf.js ${tools.pdfjs}, ${tools.chromium} (Chromium ${scored.chromium ?? 'unknown'}), fonts ${tools.fonts}`;
  const reportOf = (lane: Lane): string =>
    join(relative(process.cwd(), outOf(lane)) || '.', 'index.html');
  /** Links from the report of a lane to the reports of the other lanes of the run. */
  const links = (lane: Lane) =>
    lanes
      .filter((l) => l !== lane)
      .map((l) => ({
        label: `the ${LANE_INFO[l].title} lane`,
        href: relative(outOf(lane), join(outOf(l), 'index.html')),
      }));

  let exitCode = 0;
  for (const lane of lanes) {
    const info = LANE_INFO[lane];
    const results = targets.map((d) => byKey.get(`${lane}\0${d.name}`)!);
    const all = results.flatMap((r) => r.slides);
    const passing = all.filter((s) => s.status === 'pass').length;
    const hidden = all.filter((s) => s.hidden).length;
    const hiddenText = hidden ? `, ${hidden} hidden` : '';
    const report = reportOf(lane);
    const context = {
      tools: toolsText,
      lane: info.title,
      method: info.method,
      links: links(lane),
    };

    if (set !== 'fixtures') {
      writeReport(outOf(lane), set, results, context);
      console.log(
        `${info.title}: ${passing} of ${all.length} slides pass${hiddenText}, in ${seconds} s. Report: ${report}`
      );
      const failed = results.some(
        (r) => r.error || r.slides.some((s) => s.status === 'fail' || s.selfConsistent === false)
      );
      if (failed) exitCode = 1;
      continue;
    }

    let list: KnownGapList = laneGaps(lane);
    let refused = false;
    if (args.updateGaps) {
      // A partial run, or a deck that did not finish, replaces no entry of the decks it lacks.
      const plan = planUpdate(list, results, tools, args.acceptWorse, !only.length);
      // The regressions first: an update does not record them away.
      for (const p of checkFixtureGate(results, list, tools, !only.length).problems)
        if (p.kind === 'fix') console.log(`${info.prefix}FIX: ${p.text}`);
      for (const e of plan.changes.tools) console.log(`${info.prefix}TOOLS: ${e}`);
      for (const e of plan.changes.looser) console.log(`${info.prefix}LOOSER: ${e}`);
      for (const e of plan.changes.added) console.log(`${info.prefix}NEW: ${e}`);
      for (const e of plan.changes.removed) console.log(`${info.prefix}REMOVED: ${e}`);
      refused = !plan.write;
      if (plan.write) {
        await writeGaps(THRESHOLDS_FILE, plan.next, info.gaps);
        list = plan.next;
      }
    }
    const gate = checkFixtureGate(results, list, tools, !only.length);
    for (const r of results)
      for (const s of r.slides) {
        const key = slideKey(r.name, s.index);
        s.knownGap = gate.knownGaps.has(key);
        s.verdict = slideVerdict(s, Object.hasOwn(list.slides, key) ? list.slides[key] : undefined);
      }
    const fixes = gate.problems.filter((p) => p.kind === 'fix');
    const records = gate.problems.filter((p) => p.kind === 'record');
    const toolsStep = gate.problems.some((p) => p.kind === 'tools')
      ? 'The list was measured with other tools (TOOLS above): record it in a pull request that changes only the tools, with bun run fidelity --update-gaps --accept-worse.'
      : '';
    const step = refused
      ? `Nothing was written to ${info.gaps}: this lane has new entries or kinds (LOOSER, NEW above), and the list of another lane may have been written. Fix the regression. Only a new fixture, a pull request that changes no engine or harness source, or a pull request that changes only the comparison (COMPARISON_FILES in thresholds.ts) may record them, with: bun run fidelity <deck> --lane ${lane} --update-gaps --accept-worse`
      : args.updateGaps
        ? `${info.gaps} in scripts/fidelity/thresholds.ts now records this run. Review its diff: a new kind is a regression to fix, not to record.`
        : fixes.length
          ? `Fix the problems above; see ${report}.`
          : records.length
            ? 'The list names checks that pass now (RECORD above). Remove them with: bun run fidelity --update-gaps'
            : toolsStep
              ? ''
              : 'Nothing to do.';
    const next = [toolsStep, step].filter(Boolean).join(' ');
    const problems = gate.problems.map((p) => p.text);
    writeReport(outOf(lane), set, results, { ...context, gate: { problems, next } });
    console.log(
      `\n${info.title}: ${passing} of ${all.length} slides pass; ${gate.knownGaps.size} fail as known gaps${hiddenText}. ${seconds} s. Report: ${report}`
    );
    for (const p of gate.problems) console.log(`${info.prefix}${p.kind.toUpperCase()}: ${p.text}`);
    console.log(
      `${info.title} gate: ${gate.problems.length ? 'fails' : 'passes'}${only.length ? ` (only ${only.join(', ')})` : ''}. ${next}`
    );
    if (gate.problems.length || refused) exitCode = 1;
  }
  return exitCode;
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
  const out = resolve(args.out);
  mkdirSync(out, { recursive: true });
  const resultsFile = join(out, 'results.jsonl');
  const identity = {
    schema: CORPUS_SCHEMA,
    list: sha256(new TextEncoder().encode(listText)),
    shard,
    sources: sourcesHash(resolve('.')),
    thresholds: corpusThresholds(),
    pdfjs: PDFJS_VERSION,
    chromium: pinnedChromium(),
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
    /** Our PDF, printed with the editor's paint, against the reference PDF; pdf.js reads both. */
    method: 'pdf-vs-pdf',
    /** The version that the browser reports (`chromium` is the pinned revision). */
    chromiumBuild: null as string | null,
    startedAt: new Date().toISOString(),
    finishedAt: null as string | null,
    decks: { total: mine.length, skipped: mine.length - todo.length, scored: 0, errors: 0 },
  };
  // Through a rename, so a crash never leaves half a run.json for the next --resume.
  const writeRun = (): void => writeJsonAtomic(join(out, 'run.json'), runInfo);
  writeRun();
  await withHarnessPage(async (first, control) => {
    let page = first;
    runInfo.chromiumBuild = page.context().browser()?.version() ?? null;
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
      console.log(deckLine({ ...r, name: job.id }));
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
