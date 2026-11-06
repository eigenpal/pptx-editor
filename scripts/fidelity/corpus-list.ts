// The input and output of the harness's corpus mode (`run.ts --list`), the interface that
// `eigenpal/pptx-eval` calls for comparison A. The contract is in `CONTRIBUTING.md`
// ("Fidelity harness"). Change it only together with pptx-eval, and bump `CORPUS_SCHEMA`.

import { createHash } from 'node:crypto';
import {
  existsSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  truncateSync,
  writeFileSync,
} from 'node:fs';
import { isAbsolute, join } from 'node:path';
import { approximationTitles, otherDiagnostics, type DeckResult } from './report.ts';
import { PAGE_SIZE_TOLERANCE_PT, THRESHOLDS } from './thresholds.ts';

/** Every threshold of a corpus run, for `run.json`: the page size tolerance decides `stale`. */
export const corpusThresholds = () =>
  ({ ...THRESHOLDS, pageSizeTolerancePt: PAGE_SIZE_TOLERANCE_PT }) as const;

/**
 * The version of the `run.json` and `results.jsonl` format. 3 (6 November 2026): the editor
 * prints the deck to PDF, and pdf.js draws and reads both PDFs (`run.json` `method`). Line
 * scores count the reference lines, so `linesTotal` is the count of reference lines. The size,
 * JPEG, estimated-baseline, and skipped-line fields are gone.
 */
export const CORPUS_SCHEMA = 3;

/** One deck to score. */
export interface CorpusJob {
  /** An opaque name chosen by the caller, for example the deck's SHA-256. Never a file name. */
  readonly id: string;
  /** Absolute path of the deck. */
  readonly deck: string;
  /** Absolute path of its reference PDF, or null when it has none (every slide: no reference). */
  readonly pdf: string | null;
}

/** Allowed ids: they name the image folder, so they may hold no path separator. */
const ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/;

/**
 * Parse a job list: a JSON array, or JSON Lines (one object per line). Every path must be
 * absolute; every id must be unique and match `[A-Za-z0-9][A-Za-z0-9._-]{0,127}`.
 */
export function parseJobList(text: string): CorpusJob[] {
  const trimmed = text.trim();
  const rows: unknown[] = trimmed.startsWith('[')
    ? (JSON.parse(trimmed) as unknown[])
    : trimmed
        .split('\n')
        .filter((line) => line.trim() !== '')
        .map((line) => JSON.parse(line) as unknown);
  const seen = new Set<string>();
  return rows.map((row, i) => {
    const r = row as Partial<Record<keyof CorpusJob, unknown>>;
    const where = `job ${i + 1}`;
    if (typeof r.id !== 'string' || !ID.test(r.id)) throw new Error(`${where}: invalid id`);
    if (seen.has(r.id)) throw new Error(`${where}: duplicate id ${r.id}`);
    seen.add(r.id);
    if (typeof r.deck !== 'string' || !isAbsolute(r.deck))
      throw new Error(`${where}: deck must be an absolute path`);
    if (r.pdf !== null && r.pdf !== undefined && (typeof r.pdf !== 'string' || !isAbsolute(r.pdf)))
      throw new Error(`${where}: pdf must be an absolute path or null`);
    return { id: r.id, deck: r.deck, pdf: typeof r.pdf === 'string' ? r.pdf : null };
  });
}

/** Parse `--shard i/n` (1-based). */
export function parseShard(value: string): { index: number; count: number } {
  const m = /^(\d+)\/(\d+)$/.exec(value);
  const index = Number(m?.[1]);
  const count = Number(m?.[2]);
  if (!m || count < 1 || index < 1 || index > count)
    throw new Error(`--shard must be i/n with 1 ≤ i ≤ n, not ${value}`);
  return { index, count };
}

/** The jobs of one shard: every n-th job, in list order. */
export function shardJobs<T>(jobs: readonly T[], shard: { index: number; count: number }): T[] {
  return jobs.filter((_, i) => i % shard.count === shard.index - 1);
}

/**
 * Prepare `results.jsonl` for a run. Without `resume`, remove it. With `resume`, cut off a last
 * line that a crash left without its newline, so the next line is not glued onto it. Returns the
 * ids that already have a complete line.
 */
export function openResults(resultsFile: string, resume: boolean): Set<string> {
  if (!resume) {
    rmSync(resultsFile, { force: true });
    return new Set();
  }
  if (existsSync(resultsFile)) {
    const text = readFileSync(resultsFile, 'utf8');
    if (text && !text.endsWith('\n'))
      truncateSync(resultsFile, Buffer.byteLength(text.slice(0, text.lastIndexOf('\n') + 1)));
  }
  return finishedIds(resultsFile);
}

/**
 * Remove paths and addresses from a message before it enters the corpus output, which holds no
 * paths and no file names (`CONTRIBUTING.md`, "Fidelity harness"). Keeps the first line, and cuts it at
 * the first path: a path may hold spaces, so nothing after it is safe to keep.
 */
export function redactPaths(message: string): string {
  return (message.split('\n')[0] ?? '')
    .replace(/\b(?:https?|file):\/\/\S+/g, '<address>')
    .replace(/(^|[\s'"(=])~?\/.*$/, '$1<path>');
}

/** The ids that already have a line in `results.jsonl`, for `--resume`. */
export function finishedIds(resultsFile: string): Set<string> {
  const done = new Set<string>();
  if (!existsSync(resultsFile)) return done;
  for (const line of readFileSync(resultsFile, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    try {
      const id = (JSON.parse(line) as { id?: unknown }).id;
      if (typeof id === 'string') done.add(id);
    } catch {
      // A line cut short by a crash: that deck runs again.
    }
  }
  return done;
}

/** One `results.jsonl` line: scores only, no paths, no deck text. */
export function resultLine(job: CorpusJob, r: DeckResult): string {
  return JSON.stringify({
    id: job.id,
    error: r.error ? redactPaths(r.error) : null,
    reference: r.reference ?? 'none',
    referenceIssues: (r.referenceIssues ?? []).map(redactPaths),
    milliseconds: r.milliseconds ?? null,
    slides: r.slides.map((s) => ({
      index: s.index,
      page: s.page ?? null,
      hidden: s.hidden ?? false,
      status: s.status,
      pixelMismatch: s.pixelMismatch ?? null,
      diffBoxPt: s.diffBoxPt ?? null,
      // Family names only, as `diagnostics` give them; never line texts.
      fontsSubstituted: s.fontsSubstituted ?? null,
      lineMatch: s.lineMatch ?? null,
      linesMatched: s.linesMatched ?? null,
      linesTotal: s.linesTotal ?? null,
      // A count only; never line texts.
      extraLines: s.extraLines?.length ?? null,
      baselineDeltaPt: s.baselineDelta ?? null,
      selfConsistent: s.selfConsistent ?? null,
      cause: redactPaths(s.cause),
      approximations: approximationTitles(s.diagnostics),
      diagnostics: otherDiagnostics(s.diagnostics),
    })),
  });
}

/**
 * The sources that decide a harness result: the engine, the harness page, and the fidelity
 * scripts, with the lockfile. Tests are left out. `eigenpal/pptx-eval` hashes the same kind of
 * set; DOCX Editor's validator records an engine identity the same way.
 */
export const SOURCE_PATHS = [
  'packages/core/src',
  'packages/ooxml-core/src',
  'examples/vite/src',
  'examples/vite/harness.html',
  'examples/vite/vite.config.ts',
  'scripts/fidelity',
  'package.json',
  'tsconfig.json',
  'bun.lock',
] as const;

/**
 * SHA-256 over the relative path and content of every source file, read from disk, so an
 * uncommitted or untracked edit changes it. A missing path enters the hash as missing.
 */
export function sourcesHash(root: string, paths: readonly string[] = SOURCE_PATHS): string {
  const entries: string[] = [];
  const walk = (rel: string): void => {
    const abs = join(root, rel);
    const stat = statSync(abs, { throwIfNoEntry: false });
    if (!stat) {
      entries.push(`${rel}\0missing`);
      return;
    }
    if (stat.isDirectory()) {
      for (const name of readdirSync(abs).sort()) {
        if (name === '__tests__' || name === 'node_modules' || name.includes('.test.')) continue;
        walk(`${rel}/${name}`);
      }
      return;
    }
    entries.push(`${rel}\0${createHash('sha256').update(readFileSync(abs)).digest('hex')}`);
  };
  for (const p of paths) walk(p);
  return createHash('sha256').update(entries.join('\n')).digest('hex');
}

/** What decides whether two corpus runs may share one `results.jsonl`. */
export interface CorpusRunIdentity {
  readonly schema: number;
  readonly list: string;
  readonly shard: { readonly index: number; readonly count: number };
  readonly sources: string;
  readonly thresholds: unknown;
  readonly pdfjs: string;
  /** Playwright's pinned Chromium revision (`pinnedChromium`). */
  readonly chromium?: string;
  /** Whether the reference fonts are installed, and `localFontsIdentity` of the font files. */
  readonly referenceFonts: boolean;
  readonly fonts: string;
}

/**
 * Decide how a corpus run starts in `out`: the ids already done, or why it may not start. A run
 * without `resume` refuses an existing `results.jsonl` unless `fresh` says to replace it. A
 * resume refuses an earlier run with another list, shard, source set, thresholds, pdf.js, or Chromium revision.
 */
export function startCorpusRun(
  out: string,
  identity: CorpusRunIdentity,
  options: { readonly resume: boolean; readonly fresh: boolean }
): { readonly done: Set<string> } | { readonly refused: string } {
  const resultsFile = join(out, 'results.jsonl');
  const hasResults = existsSync(resultsFile) && statSync(resultsFile).size > 0;
  if (!options.resume) {
    if (hasResults && !options.fresh)
      return {
        refused:
          'results.jsonl exists in --out. Add --resume to continue that run, or --fresh to replace it.',
      };
    return { done: openResults(resultsFile, false) };
  }
  const runFile = join(out, 'run.json');
  if (hasResults) {
    let before: Partial<CorpusRunIdentity>;
    try {
      before = JSON.parse(readFileSync(runFile, 'utf8')) as Partial<CorpusRunIdentity>;
    } catch {
      return {
        refused:
          '--resume needs the run.json of the earlier run, and it is missing or broken. Use a new --out folder, or --fresh to replace the results.',
      };
    }
    const differ = (Object.keys(identity) as (keyof CorpusRunIdentity)[]).filter(
      (k) => JSON.stringify(before[k] ?? null) !== JSON.stringify(identity[k])
    );
    if (differ.length)
      return {
        refused: `--resume needs the same run, but these differ from the earlier run: ${differ.join(', ')}. Use a new --out folder, or --fresh to replace the results.`,
      };
  }
  return { done: openResults(resultsFile, true) };
}

/** Write JSON through a temporary file and a rename, so a crash never leaves half a file. */
export function writeJsonAtomic(path: string, value: unknown): void {
  const tmp = `${path}.${process.pid}.tmp`;
  writeFileSync(tmp, `${JSON.stringify(value, null, 2)}\n`);
  renameSync(tmp, path);
}
