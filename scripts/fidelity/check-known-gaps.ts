// The known-gap ratchet in CI, one rule: `FIXTURE_KNOWN_GAPS` takes a new entry, a new kind,
// or fewer checked lines only in a change that touches no engine or harness source
// (`nonSourceFiles`), or for a new fixture. A change to the comparison alone (`comparisonOnly`,
// with the base's `COMPARISON_FILES`) may take a new entry or kind, but not fewer checked lines.
// A removed entry is always fine. No limit in
// `thresholds.ts` loosens, and none goes away unless `RETIRED_LIMITS` gives the reason.
//
//   bun scripts/fidelity/check-known-gaps.ts --base origin/main

import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { gapKinds } from './known-gaps.ts';
import {
  FIXTURE_KNOWN_GAPS,
  PAGE_SIZE_TOLERANCE_PT,
  RETIRED_LIMITS,
  THRESHOLDS,
  type KnownGapList,
} from './thresholds.ts';
import { readManifest, type ReferenceManifest } from './reference-identity.ts';

/** The packages that are measuring tools: pdf.js, Playwright and its Chromium, the engine fonts. */
export const TOOL_PACKAGES: ReadonlySet<string> = new Set([
  'pdfjs-dist',
  'playwright',
  'playwright-core',
  '@playwright/test',
  '@docx-editor.dev/fonts',
  '@docx-editor.dev/fonts-cjk',
]);

const DEPENDENCY_FIELDS = [
  'dependencies',
  'devDependencies',
  'peerDependencies',
  'optionalDependencies',
] as const;

type Json = Record<string, unknown>;

/** A copy of a manifest object without the tool packages in its dependency fields. */
function withoutTools(manifest: Json, tools: ReadonlySet<string>): Json {
  const out: Json = { ...manifest };
  for (const field of DEPENDENCY_FIELDS) {
    const deps = manifest[field];
    if (deps && typeof deps === 'object')
      out[field] = Object.fromEntries(
        Object.entries(deps as Json).filter(([name]) => !tools.has(name))
      );
  }
  return out;
}

/** JSON with trailing commas (`bun.lock`). Undefined when it does not parse. */
function parseLoose(text: string | undefined): Json | undefined {
  if (text === undefined) return undefined;
  try {
    return JSON.parse(text.replace(/,(\s*[}\]])/g, '$1')) as Json;
  } catch {
    return undefined;
  }
}

/** Whether a `package.json` changed only in the versions of tool packages. */
export function toolOnlyManifestChange(
  base: string | undefined,
  head: string | undefined
): boolean {
  const [b, h] = [parseLoose(base), parseLoose(head)];
  if (!b || !h) return false;
  return (
    JSON.stringify(withoutTools(b, TOOL_PACKAGES)) ===
    JSON.stringify(withoutTools(h, TOOL_PACKAGES))
  );
}

/** The package name of a `bun.lock` entry: `["name@version", ...]`. */
const lockName = (entry: unknown): string | undefined => {
  const id = Array.isArray(entry) ? entry[0] : undefined;
  return typeof id === 'string' ? id.slice(0, id.lastIndexOf('@')) : undefined;
};

/** The tool packages and every package that they depend on, in one lockfile. */
function toolClosure(lock: Json): Set<string> {
  const packages = (lock.packages ?? {}) as Json;
  const byName = new Map<string, unknown>();
  for (const entry of Object.values(packages)) {
    const name = lockName(entry);
    if (name) byName.set(name, entry);
  }
  const out = new Set<string>();
  const queue = [...TOOL_PACKAGES];
  while (queue.length) {
    const name = queue.pop()!;
    if (out.has(name)) continue;
    out.add(name);
    const meta = (byName.get(name) as unknown[] | undefined)?.find(
      (part) => part && typeof part === 'object' && !Array.isArray(part)
    ) as Json | undefined;
    for (const field of DEPENDENCY_FIELDS)
      for (const dep of Object.keys((meta?.[field] as Json | undefined) ?? {})) queue.push(dep);
  }
  return out;
}

/** A lockfile without the tool packages and their dependencies. */
function lockWithoutTools(lock: Json, tools: ReadonlySet<string>): Json {
  const workspaces = Object.fromEntries(
    Object.entries((lock.workspaces ?? {}) as Json).map(([k, w]) => [
      k,
      withoutTools(w as Json, tools),
    ])
  );
  const packages = Object.fromEntries(
    Object.entries((lock.packages ?? {}) as Json).filter(([, entry]) => {
      const name = lockName(entry);
      return !name || !tools.has(name);
    })
  );
  return { ...lock, workspaces, packages };
}

/** Whether `bun.lock` changed only in the resolved versions of the tool packages. */
export function toolOnlyLockChange(base: string | undefined, head: string | undefined): boolean {
  const [b, h] = [parseLoose(base), parseLoose(head)];
  if (!b || !h) return false;
  const tools = new Set([...toolClosure(b), ...toolClosure(h)]);
  return JSON.stringify(lockWithoutTools(b, tools)) === JSON.stringify(lockWithoutTools(h, tools));
}

/** A changed file at the base and in the checkout (undefined where it does not exist). */
export type ReadChange = (file: string) => { base?: string; head?: string };

/**
 * The changed files that are engine or harness source. Everything else is measuring input:
 * `fixtures/` (decks, reference PDFs, `references.json`, fixture fonts), the fixture builder
 * `scripts/fixtures/`, `thresholds.ts`, and a `package.json` or `bun.lock` whose only change is
 * the versions of `TOOL_PACKAGES` (and of the packages they depend on).
 */
export function nonSourceFiles(files: readonly string[], read: ReadChange): string[] {
  return files.filter((f) => {
    if (
      f.startsWith('fixtures/') ||
      f.startsWith('scripts/fixtures/') ||
      f === 'scripts/fidelity/thresholds.ts'
    )
      return false;
    if (f === 'package.json' || f.endsWith('/package.json')) {
      const { base, head } = read(f);
      return !toolOnlyManifestChange(base, head);
    }
    if (f === 'bun.lock') {
      const { base, head } = read(f);
      return !toolOnlyLockChange(base, head);
    }
    return true;
  });
}

/**
 * Whether every changed source file only compares the two PDFs: a file in `comparisonFiles` (the
 * base's `COMPARISON_FILES`), a harness test, or documentation. Such a change paints and prints
 * the same pages and reads them with other checks, so the failures it finds were there before.
 */
export function comparisonOnly(
  sources: readonly string[],
  comparisonFiles: readonly string[]
): boolean {
  return (
    sources.length > 0 &&
    sources.every(
      (f) =>
        comparisonFiles.includes(f) ||
        /^scripts\/fidelity\/[^/]+\.test\.ts$/.test(f) ||
        f.endsWith('.md') ||
        f.endsWith('.mdx')
    )
  );
}

/**
 * The fixtures that `head` has and `base` does not. None when a base fixture is gone in the same
 * change: a renamed fixture is not new.
 */
export function newFixtures(base: ReferenceManifest, head: ReferenceManifest): Set<string> {
  if (Object.keys(base.references).some((n) => !Object.hasOwn(head.references, n)))
    return new Set();
  return new Set(Object.keys(head.references).filter((n) => !Object.hasOwn(base.references, n)));
}

/**
 * The ways `head` loosens `base`: new entries, new kinds, and fewer checked lines, except for
 * the fixtures in `fresh` (new fixtures). A base list in another shape (from before kinds-only
 * entries) allows any entry. `sources` lists the changed source files; when it is empty, the
 * change only measures again, and every entry is allowed. When `comparison` is true (the change
 * only compares, `comparisonOnly`), new entries and kinds are allowed, but fewer checked lines
 * are not.
 */
export function looserGaps(
  base: KnownGapList | undefined,
  head: KnownGapList,
  fresh: ReadonlySet<string>,
  sources: readonly string[],
  comparison = false
): string[] {
  const kindsOnly = base && Object.values(base.slides ?? {}).every((g) => Array.isArray(g));
  if (!base || !kindsOnly || typeof base.pdfjs !== 'string' || !sources.length) return [];
  const problems: string[] = [];
  for (const [key, gap] of comparison ? [] : Object.entries(head.slides)) {
    if (fresh.has(key.slice(0, key.lastIndexOf('/')))) continue;
    const old = Object.hasOwn(base.slides, key) ? base.slides[key] : undefined;
    if (!old) {
      problems.push(`${key}: a new entry (${gap.join(', ')})`);
      continue;
    }
    for (const kind of gapKinds(gap))
      if (!old.includes(kind)) problems.push(`${key}: a new kind, ${kind}`);
  }
  for (const [deck, lines] of Object.entries(head.deckLines ?? {})) {
    const was = base.deckLines?.[deck];
    if (!fresh.has(deck) && was !== undefined && lines < was)
      problems.push(`${deck}: checks fewer lines (${was} → ${lines})`);
  }
  if (problems.length && !comparison) {
    const shown = sources.slice(0, 5).join(', ');
    const more = sources.length > 5 ? ` and ${sources.length - 5} more` : '';
    problems.push(
      `the change also touches source (${shown}${more}): record new failures in a pull request that changes only tools, fixtures, or references`
    );
  }
  return problems;
}

type Limits = Readonly<Record<string, number>>;

/**
 * The limits that are looser than the base's, or that the base has and `head` does not (unless
 * `retired` gives the reason): thresholds, page tolerance, and gap margins.
 */
export function looserLimits(
  base: Limits,
  head: Limits,
  retired: Readonly<Record<string, string>> = RETIRED_LIMITS
): string[] {
  const out: string[] = [];
  for (const key of Object.keys(base))
    if (!(key in head) && !Object.hasOwn(retired, key))
      out.push(
        `${key} was removed (it was ${base[key]}); list it in RETIRED_LIMITS with the reason`
      );
  for (const [key, value] of Object.entries(head)) {
    if (!(key in base)) continue;
    // A smaller raster or a smaller line share is looser; for every other limit, a larger value.
    const lowerIsLooser = key === 'renderWidthPx' || key === 'minLineBreakMatch';
    if (lowerIsLooser ? value < base[key] : value > base[key])
      out.push(`${key} is looser (${base[key]} → ${value})`);
  }
  return out;
}

/** The limits of a thresholds module as one flat record. */
const limitsOf = (m: {
  THRESHOLDS?: Record<string, unknown>;
  PAGE_SIZE_TOLERANCE_PT?: number;
  GAP_MARGINS?: Record<string, number>;
}): Limits => {
  const out: Record<string, number> = {};
  for (const [k, v] of Object.entries(m.THRESHOLDS ?? {})) if (typeof v === 'number') out[k] = v;
  if (m.PAGE_SIZE_TOLERANCE_PT !== undefined) out.pageSizeTolerancePt = m.PAGE_SIZE_TOLERANCE_PT;
  for (const [k, v] of Object.entries(m.GAP_MARGINS ?? {})) out[`margin ${k}`] = v;
  return out;
};

/** The files that differ between the merge base with `base` and the working tree. */
function changedFiles(base: string): string[] {
  try {
    const mergeBase = execFileSync('git', ['merge-base', base, 'HEAD'], {
      encoding: 'utf8',
    }).trim();
    const tracked = execFileSync('git', ['diff', '--name-only', mergeBase], { encoding: 'utf8' });
    // Files that are not added yet count too, so a local check sees the whole change.
    const untracked = execFileSync('git', ['ls-files', '--others', '--exclude-standard'], {
      encoding: 'utf8',
    });
    return `${tracked}\n${untracked}`.split('\n').filter(Boolean);
  } catch {
    return ['(the file list is not available)'];
  }
}

function gitShow(ref: string, path: string): string | undefined {
  try {
    return execFileSync('git', ['show', `${ref}:${path}`], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    });
  } catch {
    return undefined;
  }
}

/** Check this checkout against a base ref. Returns the lines to print and the exit code. */
export async function checkAgainstBase(
  base: string
): Promise<{ lines: string[]; exitCode: number }> {
  try {
    execFileSync('git', ['rev-parse', '--verify', '--quiet', `${base}^{commit}`], {
      stdio: 'ignore',
    });
  } catch {
    return { lines: [`${base} is not a commit. Fetch the base branch first.`], exitCode: 2 };
  }
  const source = gitShow(base, 'scripts/fidelity/thresholds.ts');
  let mod: Parameters<typeof limitsOf>[0] & {
    FIXTURE_KNOWN_GAPS?: KnownGapList;
    RETIRED_LIMITS?: Record<string, string>;
    COMPARISON_FILES?: readonly string[];
  } = {};
  if (source) {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-gaps-'));
    try {
      writeFileSync(join(dir, 'thresholds.ts'), source);
      mod = await import(join(dir, 'thresholds.ts'));
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }
  const manifest = gitShow(base, 'fixtures/references.json');
  const baseList = mod.FIXTURE_KNOWN_GAPS;
  const read: ReadChange = (file) => ({
    base: gitShow(base, file),
    head: existsSync(file) ? readFileSync(file, 'utf8') : undefined,
  });
  const notices = Object.entries(RETIRED_LIMITS)
    .filter(([key]) => !Object.hasOwn(mod.RETIRED_LIMITS ?? {}, key))
    .map(([key, reason]) => `NOTICE: ${key} is retired: ${reason}`);
  const sources = nonSourceFiles(changedFiles(base), read);
  // The base's list, so a change cannot widen the exception for itself.
  const comparison = comparisonOnly(sources, mod.COMPARISON_FILES ?? []);
  const problems = [
    ...looserLimits(limitsOf(mod), limitsOf({ THRESHOLDS, PAGE_SIZE_TOLERANCE_PT })),
    ...looserGaps(
      baseList,
      FIXTURE_KNOWN_GAPS,
      newFixtures(
        manifest ? (JSON.parse(manifest) as ReferenceManifest) : { version: 1, references: {} },
        readManifest('fixtures')
      ),
      sources,
      comparison
    ),
  ];
  if (!problems.length)
    return {
      lines: [...notices, `Nothing in the known gaps or the limits loosens against ${base}.`],
      exitCode: 0,
    };
  return {
    lines: [
      ...notices,
      ...problems.map((p) => `GAPS: ${p}`),
      'The known gaps and the limits only tighten, and no limit goes away unless RETIRED_LIMITS gives the reason. Fix the regression instead of recording it. Only a new fixture, a pull request that changes no engine or harness source, or a pull request that changes only the comparison (COMPARISON_FILES) may take a new entry or kind.',
    ],
    exitCode: 1,
  };
}

if (import.meta.main) {
  let base: string | undefined;
  try {
    ({
      values: { base },
    } = parseArgs({ strict: true, options: { base: { type: 'string' } } }));
  } catch (e) {
    console.error((e as Error).message);
  }
  if (!base) {
    console.error('Usage: bun scripts/fidelity/check-known-gaps.ts --base <git ref>');
    process.exit(2);
  }
  const result = await checkAgainstBase(base);
  for (const line of result.lines) console.log(line);
  process.exit(result.exitCode);
}
