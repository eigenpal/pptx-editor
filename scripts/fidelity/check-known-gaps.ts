// The known-gap ratchet in CI: `FIXTURE_KNOWN_GAPS` only shrinks, and no limit in
// `thresholds.ts` loosens. `bun run fidelity --update-gaps` writes whatever a run scores, so this
// check compares the list with the base branch. A new entry, a new kind, or a looser value is
// allowed only for a fixture whose deck or reference PDF changed in `references.json`, or when the
// list was measured with other tools (pdf.js, Chromium, fixture fonts).
//
//   bun scripts/fidelity/check-known-gaps.ts --base origin/main

import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { gapKinds } from './known-gaps.ts';
import {
  FIXTURE_KNOWN_GAPS,
  GAP_MARGINS,
  PAGE_SIZE_TOLERANCE_PT,
  THRESHOLDS,
  type FailureKind,
  type KnownGap,
  type KnownGapList,
} from './thresholds.ts';
import { readManifest, type ReferenceManifest } from './reference-identity.ts';

/** Whether `head` allows a worse score than `base` for one kind. */
function looser(kind: FailureKind, head: KnownGap, base: KnownGap): boolean {
  if (kind === 'pixels') return (head.pixels ?? 0) > (base.pixels ?? 0);
  if (kind === 'line-breaks') return (head['line-breaks'] ?? 1) < (base['line-breaks'] ?? 1);
  return (head.baseline ?? 0) > (base.baseline ?? 0);
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
 * or measured with other tools allows any value: the scores are new measurements.
 */
export function looserGaps(
  base: KnownGapList | undefined,
  head: KnownGapList,
  changed: ReadonlySet<string>
): string[] {
  const tools = ['pdfjs', 'chromium', 'fonts'] as const;
  if (!base || typeof base.pdfjs !== 'string' || tools.some((k) => base[k] !== head[k])) return [];
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
  return problems;
}

type Limits = Readonly<Record<string, number>>;

/** The limits that are looser than the base's: thresholds, page tolerance, and gap margins. */
export function looserLimits(base: Limits, head: Limits): string[] {
  const out: string[] = [];
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
  const problems = [
    ...looserLimits(limitsOf(mod), limitsOf({ THRESHOLDS, PAGE_SIZE_TOLERANCE_PT, GAP_MARGINS })),
    ...looserGaps(
      mod.FIXTURE_KNOWN_GAPS,
      FIXTURE_KNOWN_GAPS,
      changedReferences(
        manifest ? (JSON.parse(manifest) as ReferenceManifest) : { version: 1, references: {} },
        readManifest('fixtures')
      )
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
      'The known gaps and the limits only tighten. Fix the regression instead of recording it. Only a new fixture, a changed deck, a new reference PDF, or new tools may take a new or looser entry.',
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
