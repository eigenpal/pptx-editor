// The known-gap ratchet in CI: `FIXTURE_KNOWN_GAPS` only shrinks, and no limit in
// `thresholds.ts` loosens or goes away. This check compares the list with the base branch. A new
// entry, a new kind, or a looser value is allowed only for a fixture whose deck or reference PDF
// changed in `references.json`, or when the list was measured with other tools (pdf.js,
// Chromium, fixture fonts) and the change touches nothing but those tools: the lockfile,
// `package.json` files, `thresholds.ts`, and, for new fixture fonts, `fixtures/fonts/`.
//
//   bun scripts/fidelity/check-known-gaps.ts --base origin/main

import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { gapKinds, looser } from './known-gaps.ts';
import {
  FIXTURE_KNOWN_GAPS,
  GAP_MARGINS,
  PAGE_SIZE_TOLERANCE_PT,
  RETIRED_LIMITS,
  THRESHOLDS,
  type KnownGapList,
} from './thresholds.ts';
import { readManifest, type ReferenceManifest } from './reference-identity.ts';

const TOOLS = ['pdfjs', 'chromium', 'fonts'] as const;

/**
 * The changed files that are not tool files: everything but the lockfile, `package.json` files,
 * `thresholds.ts`, and (when the fixture fonts changed) `fixtures/fonts/`.
 */
export function nonToolFiles(files: readonly string[], fontsChanged: boolean): string[] {
  return files.filter(
    (f) =>
      f !== 'bun.lock' &&
      f !== 'package.json' &&
      !f.endsWith('/package.json') &&
      f !== 'scripts/fidelity/thresholds.ts' &&
      !(fontsChanged && f.startsWith('fixtures/fonts/'))
  );
}

/** The fixtures whose deck or reference PDF differs between two manifests, or that are new. */
export function changedReferences(base: ReferenceManifest, head: ReferenceManifest): Set<string> {
  const changed = new Set<string>();
  for (const [name, h] of Object.entries(head.references)) {
    const b = Object.hasOwn(base.references, name) ? base.references[name] : undefined;
    if (!b || b.deck !== h.deck || b.pdf !== h.pdf) changed.add(name);
  }
  return changed;
}

/**
 * The ways `head` loosens `base`. A base list in another shape (from before the PDF comparison)
 * allows any value. A list measured with other tools allows any value only when the change
 * touches nothing but the tools (`otherFiles` is empty); else every looser entry is listed, with
 * the files that keep it from counting as a tools change.
 */
export function looserGaps(
  base: KnownGapList | undefined,
  head: KnownGapList,
  changed: ReadonlySet<string>,
  otherFiles: readonly string[] = []
): string[] {
  if (!base || typeof base.pdfjs !== 'string') return [];
  const newTools = TOOLS.filter((k) => base[k] !== head[k]);
  if (newTools.length && !otherFiles.length) return [];
  const problems: string[] = [];
  for (const [key, gap] of Object.entries(head.slides)) {
    if (changed.has(key.slice(0, key.lastIndexOf('/')))) continue;
    const old = Object.hasOwn(base.slides, key) ? base.slides[key] : undefined;
    if (!old) {
      problems.push(`${key}: a new entry, but the fixture's reference did not change`);
      continue;
    }
    if ((gap.lines ?? 0) < (old.lines ?? 0))
      problems.push(`${key}: fewer checked lines (${old.lines} → ${gap.lines ?? 0})`);
    for (const kind of gapKinds(gap))
      if (old[kind] === undefined) problems.push(`${key}: a new kind, ${kind}`);
      else if (looser(kind, gap, old))
        problems.push(`${key}: a looser ${kind} score (${old[kind]} → ${gap[kind]})`);
  }
  if (problems.length && newTools.length) {
    const shown = otherFiles.slice(0, 5).join(', ');
    const more = otherFiles.length > 5 ? ` and ${otherFiles.length - 5} more` : '';
    problems.push(
      `the list was measured with new tools (${newTools.join(', ')}), but the change also touches ${shown}${more}: move the tools change to its own pull request`
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
    return execFileSync('git', ['diff', '--name-only', mergeBase], { encoding: 'utf8' })
      .split('\n')
      .filter(Boolean);
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
  let mod: Parameters<typeof limitsOf>[0] & { FIXTURE_KNOWN_GAPS?: KnownGapList } = {};
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
  const fontsChanged = baseList !== undefined && baseList.fonts !== FIXTURE_KNOWN_GAPS.fonts;
  const problems = [
    ...looserLimits(limitsOf(mod), limitsOf({ THRESHOLDS, PAGE_SIZE_TOLERANCE_PT, GAP_MARGINS })),
    ...looserGaps(
      baseList,
      FIXTURE_KNOWN_GAPS,
      changedReferences(
        manifest ? (JSON.parse(manifest) as ReferenceManifest) : { version: 1, references: {} },
        readManifest('fixtures')
      ),
      nonToolFiles(changedFiles(base), fontsChanged)
    ),
  ];
  if (!problems.length)
    return {
      lines: [`Nothing in the known gaps or the limits loosens against ${base}.`],
      exitCode: 0,
    };
  return {
    lines: [
      ...problems.map((p) => `GAPS: ${p}`),
      'The known gaps and the limits only tighten, and no limit goes away. Fix the regression instead of recording it. Only a new fixture, a changed deck, a new reference PDF, or a pull request that changes only the tools may take a new or looser entry.',
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
