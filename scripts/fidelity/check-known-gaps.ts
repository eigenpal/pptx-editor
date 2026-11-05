// The known-gap ratchet in CI:`FIXTURE_KNOWN_GAPS` only shrinks, and no limit in
// `thresholds.ts` loosens. Compared with the base branch, a change may not add an entry, add a
// kind, or loosen a value, unless the fixture's reference really changed in the same change: a
// new fixture, a changed deck with its new reference PDF, or a new reference PDF of the same
// deck. A list that records new Poppler or Chromium versions may take new values (the change
// measured it again with those tools). The one move to supersampled rasters
// (`RASTER_SCALE_MOVE`) and the one move to anti-aliasing judged at the device scale
// (`ANTIALIASING_MOVE`) each allow new values once; any other change of either is a problem. One
// looser value has an owner exception (`OWNER_GAP_EXCEPTION`), for exactly one slide. A new list is allowed only together with its new environment in
// `FIDELITY_ENVIRONMENTS`, and the check covers every list; a removed list is a problem.
//
//   bun run fidelity:check-gaps -- --base origin/main

import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { gapKinds } from './known-gaps.ts';
import {
  FIDELITY_ENVIRONMENTS,
  FIXTURE_KNOWN_GAPS,
  GAP_MARGINS,
  PAGE_SIZE_TOLERANCE_PT,
  THRESHOLDS,
  type FailureKind,
  type KnownGap,
} from './thresholds.ts';
import {
  readManifest,
  type ReferenceIdentity,
  type ReferenceManifest,
} from './reference-identity.ts';

type Entries = Readonly<Record<string, KnownGap | readonly FailureKind[]>>;
/** The lists by environment: `{ poppler, chromium, slides }`, or (first shape) the entries alone. */
type Lists = Readonly<Record<string, Entries | { readonly slides: Entries }>>;
type Thresholds = Readonly<Record<string, number | string>>;

/** The entries of one environment's list, in either shape. */
const entriesOf = (list: Entries | { readonly slides: Entries }): Entries =>
  'slides' in list && typeof list.slides === 'object' && !Array.isArray(list.slides)
    ? (list.slides as Entries)
    : (list as Entries);

/** A base entry in the first, kinds-only shape has no score: any score of its kinds is allowed. */
function normalize(gap: KnownGap | readonly FailureKind[]): KnownGap {
  if (!Array.isArray(gap)) return gap as KnownGap;
  const kinds = gap as readonly FailureKind[];
  return {
    size: kinds.includes('size') ? true : undefined,
    pixels: kinds.includes('pixels') ? Infinity : undefined,
    'line-breaks': kinds.includes('line-breaks') ? -Infinity : undefined,
    baseline: kinds.includes('baseline') ? Infinity : undefined,
  };
}

/** Whether `head` allows a worse score than `base` for one kind. */
function looser(kind: FailureKind, head: KnownGap, base: KnownGap): boolean {
  if (kind === 'pixels') return (head.pixels ?? 0) > (base.pixels ?? 0);
  if (kind === 'line-breaks') return (head['line-breaks'] ?? 1) < (base['line-breaks'] ?? 1);
  if (kind === 'baseline') return (head.baseline ?? 0) > (base.baseline ?? 0);
  return false;
}

/**
 * The one change of `THRESHOLDS.rasterScale` that may re-record the known gaps: from scale 1
 * (no supersampling) to scale 4 (`supersample.ts`). It happened once, in the change that added
 * supersampling. Another scale does not make the comparison more exact (scale 8 breaks a pattern
 * fill, `CONTRIBUTING.md`, "Supersampled rasters"), so a later scale change could only be a way
 * to record other values. It is a problem, and so is any list that takes new values from it.
 */
export const RASTER_SCALE_MOVE = { from: 1, to: 4 } as const;

/** Whether a raster scale change from `base` to `head` is the one allowed move. */
const allowedScaleMove = (base: unknown, head: unknown): boolean =>
  base === RASTER_SCALE_MOVE.from && head === RASTER_SCALE_MOVE.to;

/**
 * The one change of `THRESHOLDS.antialiasing` that may re-record the known gaps: from judging
 * anti-aliasing on the compared images (`images`) to judging it on the device-scale images
 * (`device`, `supersample.ts`). On box-filtered images, pixelmatch called a stroke moved by 1 px
 * anti-aliasing, so the `device` mode is stricter. There is no other mode to move to: any other
 * change is a problem, and so is any list that takes new values from it. A list or a
 * `THRESHOLDS` without the field judged anti-aliasing on the compared images.
 */
export const ANTIALIASING_MOVE = { from: 'images', to: 'device' } as const;

/**
 * The one looser known-gap value that the owner allowed (5 November 2026, eigenpal/pptx-editor-miro#156,
 * whole-point baselines): `rtl-text/4` on `linux` scores 3.41 % of pixels instead of 3.40 %. Its
 * geometry is better: each of its 12 lines is within 0.12 pt of its PDF glyph origin, against up
 * to 0.49 pt without the rule. The pixels rise because CI draws the slide with substitute glyph
 * outlines and no line breaks at the same words as the PDF (`line-breaks` 0). It allows exactly
 * this slide, environment, kind, and pair of values. Once the base records `to`, it allows nothing.
 */
export const OWNER_GAP_EXCEPTION = {
  environment: 'linux',
  slide: 'rtl-text/4',
  kind: 'pixels',
  from: 0.034,
  to: 0.0341,
} as const;

/** Whether a looser `kind` score of `key` in `env` is the one owner exception. */
const ownerException = (
  env: string,
  key: string,
  kind: string,
  base: unknown,
  head: unknown
): boolean =>
  env === OWNER_GAP_EXCEPTION.environment &&
  key === OWNER_GAP_EXCEPTION.slide &&
  kind === OWNER_GAP_EXCEPTION.kind &&
  base === OWNER_GAP_EXCEPTION.from &&
  head === OWNER_GAP_EXCEPTION.to;

/** The anti-aliasing mode of a list or of `THRESHOLDS`; `images` when it has none. */
const antialiasingOf = (values: Readonly<Record<string, unknown>>): unknown =>
  values.antialiasing ?? ANTIALIASING_MOVE.from;

/** Whether an anti-aliasing mode change from `base` to `head` is the one allowed move. */
const allowedAntialiasingMove = (base: unknown, head: unknown): boolean =>
  base === ANTIALIASING_MOVE.from && head === ANTIALIASING_MOVE.to;

/**
 * Why a list may take new values: it records other Poppler or Chromium versions than the base's,
 * because the change measured it again with those tools, or it made one of the two one-time
 * moves of the comparison: from raster scale 1 to 4 (`RASTER_SCALE_MOVE`), or from anti-aliasing
 * judged on the compared images to the device scale (`ANTIALIASING_MOVE`). A list with tools but
 * no `rasterScale` was measured at scale 1, and one without `antialiasing` judged it on the
 * compared images. Any other change of either allows nothing. Undefined when nothing allows new
 * values.
 */
export function toolChange(
  base: Entries | { readonly slides: Entries },
  head: Entries | { readonly slides: Entries }
): string | undefined {
  const tools = (list: object): Record<string, unknown> =>
    'slides' in list ? (list as Record<string, unknown>) : {};
  const b = tools(base);
  const h = tools(head);
  const changed = (['poppler', 'chromium'] as const)
    .filter((k) => b[k] !== undefined && h[k] !== undefined && b[k] !== h[k])
    .map(
      (k) => `${k === 'poppler' ? 'Poppler' : 'Chromium'} ${String(h[k])} (was ${String(b[k])})`
    );
  const scale = (list: Record<string, unknown>): unknown =>
    list.slides === undefined ? undefined : (list.rasterScale ?? 1);
  if (scale(b) !== undefined && scale(h) !== undefined && allowedScaleMove(scale(b), scale(h)))
    changed.push(`raster scale ${String(scale(h))} (was ${String(scale(b))})`);
  if (
    b.slides !== undefined &&
    h.slides !== undefined &&
    allowedAntialiasingMove(antialiasingOf(b), antialiasingOf(h))
  )
    changed.push(
      `anti-aliasing judged on the ${String(antialiasingOf(h))} images (was ${String(antialiasingOf(b))})`
    );
  return changed.length ? `the list was measured again with ${changed.join(', ')}` : undefined;
}

/** How each fixture of `head` relates to `base`: its base name, and why it may change. */
export interface ReferenceChanges {
  /** The fixture's name in the base, when it was renamed (same deck bytes, other name). */
  readonly renamed: ReadonlyMap<string, string>;
  /** Fixtures whose reference really changed, with the reason. */
  readonly changed: ReadonlyMap<string, string>;
  /**
   * The changed fixtures whose own line count may change: a new fixture or a changed deck. A new
   * PDF of the same deck does not change the editor's lines.
   */
  readonly content?: ReadonlySet<string>;
}

/**
 * The fixtures whose reference is new or really changed between two manifests.
 *
 * The manifest records only the deck and PDF hashes, not how a PDF was made. So a new PDF hash
 * for the same deck hash counts as a new reference: the change replaced the reference PDF, and
 * reviewers see the reason on an ALLOWED line. A new deck hash with the old PDF hash is no change:
 * that reference is stale, and the harness reports it as a gap.
 */
export function referenceChanges(
  base: ReferenceManifest,
  head: ReferenceManifest
): ReferenceChanges {
  const renamed = new Map<string, string>();
  const changed = new Map<string, string>();
  const content = new Set<string>();
  const byDeck = new Map<string, string>();
  for (const [name, b] of Object.entries(base.references)) byDeck.set(b.deck, name);
  for (const [name, h] of Object.entries(head.references)) {
    let b: ReferenceIdentity | undefined = Object.hasOwn(base.references, name)
      ? base.references[name]
      : undefined;
    if (!b && byDeck.has(h.deck)) {
      renamed.set(name, byDeck.get(h.deck)!);
      b = base.references[byDeck.get(h.deck)!];
    }
    if (!b) {
      changed.set(name, 'a new fixture');
      content.add(name);
    } else if (b.deck !== h.deck && b.pdf !== h.pdf) {
      changed.set(name, 'its deck changed and has a new reference PDF');
      content.add(name);
    } else if (b.deck === h.deck && b.pdf !== h.pdf) {
      changed.set(name, 'its reference PDF was replaced');
    }
  }
  return { renamed, changed, content };
}

export interface GapCheck {
  /** Ways the change loosens the list. */
  readonly problems: string[];
  /** Changes that are allowed only because the fixture's reference changed, for reviewers. */
  readonly allowed: string[];
  /**
   * Removed entries that recorded a line count, for reviewers. A slide that passes now and a
   * slide whose text is gone both leave the list, and only the run's report tells them apart.
   */
  readonly removed: string[];
}

/** The environments of each side, for the rule on new lists. */
export interface Environments {
  readonly base: readonly string[];
  readonly head: readonly string[];
}

/**
 * The ways `head` loosens `base`: new entries, new kinds, and looser values, for fixtures whose
 * reference did not really change, and lists that come without their environment. Without a
 * base list (`undefined`), nothing is looser.
 */
export function looserGaps(
  base: Lists | undefined,
  head: Lists,
  references: ReferenceChanges = { renamed: new Map(), changed: new Map() },
  environments?: Environments
): GapCheck {
  const problems: string[] = [];
  const allowed: string[] = [];
  const removed: string[] = [];
  for (const env of Object.keys(head))
    if (environments && !environments.head.includes(env))
      problems.push(`${env}: a list for an environment that FIDELITY_ENVIRONMENTS does not name`);
  if (!base) return { problems, allowed, removed };
  // A removed list stops gating its environment, so it is never a quiet change.
  for (const env of Object.keys(base))
    if (!Object.hasOwn(head, env)) problems.push(`${env}: the list was removed`);
  for (const [env, list] of Object.entries(head)) {
    if (!Object.hasOwn(base, env)) {
      // A new list is allowed only in the change that adds its environment.
      if (environments && environments.base.includes(env))
        problems.push(
          `${env}: a new list for an environment that the base already has; add a list only with its new environment`
        );
      else allowed.push(`${env}: a new list, with its new environment`);
      continue;
    }
    const baseList = entriesOf(base[env]);
    const tools = toolChange(base[env], list);
    const kept = new Set<string>();
    for (const [key, raw] of Object.entries(entriesOf(list))) {
      const slash = key.lastIndexOf('/');
      const deck = key.slice(0, slash);
      const baseKey = references.renamed.has(deck)
        ? `${references.renamed.get(deck)}${key.slice(slash)}`
        : key;
      kept.add(baseKey);
      const gap = normalize(raw);
      const found: string[] = [];
      // The line count counts the editor's own lines, so no tool change explains fewer of them.
      let fewerLines: string | undefined;
      if (!Object.hasOwn(baseList, baseKey)) found.push('a new entry');
      else {
        const old = normalize(baseList[baseKey]);
        // `gapEntry` keeps `lines` on every entry whose slide checks lines, also one that drops
        // its line kinds, so a smaller or missing count means lost text.
        // A size entry records no line count, so its count is unknown, not smaller.
        if (!gap.size && old.lines !== undefined && (gap.lines ?? 0) < old.lines)
          fewerLines = `fewer checked lines (${old.lines} → ${gap.lines ?? 0})`;
        for (const kind of gapKinds(gap))
          if (old[kind] === undefined && old.size && !gap.size)
            // A size entry records no other score, so the scores that replace it are not new.
            allowed.push(
              `${env} ${key}: ${kind} replaces its size entry, which recorded no other score`
            );
          else if (old[kind] === undefined) found.push(`a new kind, ${kind}`);
          else if (looser(kind, gap, old) && ownerException(env, key, kind, old[kind], gap[kind]))
            allowed.push(
              `${env} ${key}: a looser ${kind} score (${String(old[kind])} → ${String(gap[kind])}), allowed once by the owner (OWNER_GAP_EXCEPTION)`
            );
          else if (looser(kind, gap, old))
            found.push(`a looser ${kind} score (${String(old[kind])} → ${String(gap[kind])})`);
      }
      const changed = references.changed.get(deck);
      const reason = changed ?? tools;
      for (const [f, why] of [
        ...found.map((f) => [f, reason] as const),
        // Fewer lines need a change of the deck or the fonts, not of the tools.
        ...(fewerLines
          ? [[fewerLines, references.content?.has(deck) ? changed : undefined] as const]
          : []),
      ])
        if (why) allowed.push(`${env} ${key}: ${f}, allowed because ${why}`);
        else problems.push(`${env} ${key}: ${f}, but the fixture's reference did not change`);
    }
    for (const [key, raw] of Object.entries(baseList)) {
      const lines = normalize(raw).lines;
      if (!kept.has(key) && lines)
        removed.push(
          `${env} ${key}: removed with its ${lines} checked ${lines === 1 ? 'line' : 'lines'}. If the slide still exists, check in this change's fidelity report that it passes with all ${lines === 1 ? 'of it' : `${lines} lines`}`
        );
    }
  }
  return { problems, allowed, removed };
}

/** The `THRESHOLDS` values that are looser than the base's. They only tighten. */
export function looserThresholds(base: Thresholds | undefined, head: Thresholds): string[] {
  if (!base) return [];
  const out: string[] = [];
  // A smaller raster hides small differences, so the width only grows.
  if ('renderWidthPx' in base && head.renderWidthPx < base.renderWidthPx)
    out.push(`THRESHOLDS.renderWidthPx is looser (${base.renderWidthPx} → ${head.renderWidthPx})`);
  // The raster scale changes once, from 1 to 4 (`RASTER_SCALE_MOVE`). A lower scale brings back
  // the glyph snapping bias (`supersample.ts`), and a higher one is not more exact. A base
  // without it compared at scale 1.
  const [baseScale, headScale] = [base.rasterScale ?? 1, head.rasterScale ?? 1];
  if (baseScale !== headScale && !allowedScaleMove(baseScale, headScale))
    out.push(
      `THRESHOLDS.rasterScale changed (${baseScale} → ${headScale}); only the one move from ${RASTER_SCALE_MOVE.from} to ${RASTER_SCALE_MOVE.to} is allowed`
    );
  // The anti-aliasing mode changes once, to the stricter `device` (`ANTIALIASING_MOVE`).
  const [baseAa, headAa] = [antialiasingOf(base), antialiasingOf(head)];
  if (baseAa !== headAa && !allowedAntialiasingMove(baseAa, headAa))
    out.push(
      `THRESHOLDS.antialiasing changed (${String(baseAa)} → ${String(headAa)}); only the one move from ${ANTIALIASING_MOVE.from} to ${ANTIALIASING_MOVE.to} is allowed`
    );
  // Higher is looser for these; for `minLineBreakMatch`, lower is looser.
  const higherIsLooser = [
    'maxSizeDeltaPx',
    'maxPixelMismatch',
    'pixelThreshold',
    'maxBaselineDeltaPt',
    'maxSelfConsistencyMismatch',
  ];
  for (const key of higherIsLooser)
    if (key in base && key in head && head[key] > base[key])
      out.push(`THRESHOLDS.${key} is looser (${base[key]} → ${head[key]})`);
  if (head.minLineBreakMatch < base.minLineBreakMatch)
    out.push(
      `THRESHOLDS.minLineBreakMatch is looser (${base.minLineBreakMatch} → ${head.minLineBreakMatch})`
    );
  return out;
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

/** The other limits in `thresholds.ts` that only tighten: the page-size tolerance and the margins. */
export function looserLimits(
  base: { readonly page?: number; readonly margins?: Readonly<Record<string, number>> },
  head: { readonly page: number; readonly margins: Readonly<Record<string, number>> }
): string[] {
  const out: string[] = [];
  if (base.page !== undefined && head.page > base.page)
    out.push(`PAGE_SIZE_TOLERANCE_PT is looser (${base.page} → ${head.page})`);
  for (const [k, v] of Object.entries(head.margins))
    if (base.margins && k in base.margins && v > base.margins[k])
      out.push(`GAP_MARGINS.${k} is looser (${base.margins[k]} → ${v})`);
  return out;
}

interface BaseValues {
  readonly lists?: Lists;
  readonly thresholds?: Thresholds;
  readonly environments?: readonly string[];
  readonly page?: number;
  readonly margins?: Readonly<Record<string, number>>;
}

/** The base branch's lists and limits; each is undefined when the base has none yet. */
async function baseValues(ref: string): Promise<BaseValues> {
  const source = gitShow(ref, 'scripts/fidelity/thresholds.ts');
  if (!source) return {};
  const dir = mkdtempSync(join(tmpdir(), 'pptx-gaps-'));
  try {
    const file = join(dir, 'thresholds.ts');
    writeFileSync(file, source);
    const mod = (await import(file)) as {
      FIXTURE_KNOWN_GAPS?: Lists;
      THRESHOLDS?: Thresholds;
      FIDELITY_ENVIRONMENTS?: readonly string[];
      PAGE_SIZE_TOLERANCE_PT?: number;
      GAP_MARGINS?: Readonly<Record<string, number>>;
    };
    return {
      lists: mod.FIXTURE_KNOWN_GAPS,
      thresholds: mod.THRESHOLDS,
      // Before the constant, the base's environments are the ones its lists name.
      environments:
        mod.FIDELITY_ENVIRONMENTS ??
        (mod.FIXTURE_KNOWN_GAPS ? Object.keys(mod.FIXTURE_KNOWN_GAPS) : []),
      page: mod.PAGE_SIZE_TOLERANCE_PT,
      margins: mod.GAP_MARGINS,
    };
  } finally {
    rmSync(dir, { recursive: true, force: true });
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
  const baseManifestText = gitShow(base, 'fixtures/references.json');
  const baseManifest: ReferenceManifest = baseManifestText
    ? (JSON.parse(baseManifestText) as ReferenceManifest)
    : { version: 1, references: {} };
  const values = await baseValues(base);
  const gaps = looserGaps(
    values.lists,
    FIXTURE_KNOWN_GAPS as unknown as Lists,
    referenceChanges(baseManifest, readManifest('fixtures')),
    { base: values.environments ?? [], head: FIDELITY_ENVIRONMENTS }
  );
  const problems = [
    ...looserThresholds(values.thresholds, THRESHOLDS as unknown as Thresholds),
    ...looserLimits(values, { page: PAGE_SIZE_TOLERANCE_PT, margins: GAP_MARGINS }),
    ...gaps.problems,
  ];
  const lines = [
    ...(values.lists
      ? []
      : [
          `${base} has no known-gap list, so every entry here is new; only the limits are compared.`,
        ]),
    ...gaps.allowed.map((a) => `ALLOWED: ${a}`),
    ...gaps.removed.map((r) => `NOTE: ${r}`),
    ...problems.map((p) => `GAPS: ${p}`),
  ];
  if (problems.length) {
    lines.push(
      'The known gaps and the thresholds only tighten. Undo each change that a GAPS line names. Only a new fixture, a changed deck with its new PDF, or a new reference PDF of the same deck may add an entry, a kind, or a looser value; no change loosens a limit.'
    );
    return { lines, exitCode: 1 };
  }
  lines.push(
    values.lists
      ? `Nothing is added or loosened against ${base}, except what ALLOWED lines name.`
      : `No limit is loosened against ${base}.`
  );
  return { lines, exitCode: 0 };
}

export const CHECK_GAPS_USAGE = `Usage: bun run fidelity:check-gaps -- --base <git ref, for example origin/main>
  Check that FIXTURE_KNOWN_GAPS and THRESHOLDS only tighten against the base branch.

Exit codes: 0 nothing loosens, 1 something loosens, 2 a usage error or an unknown ref.`;

if (import.meta.main) {
  let base: string | undefined;
  let help: boolean | undefined;
  try {
    ({
      values: { base, help },
    } = parseArgs({
      strict: true,
      options: { base: { type: 'string' }, help: { type: 'boolean', short: 'h' } },
    }));
  } catch (e) {
    console.error(`fidelity:check-gaps: ${(e as Error).message}\n\n${CHECK_GAPS_USAGE}`);
    process.exit(2);
  }
  if (help || !base) {
    (help ? console.log : console.error)(CHECK_GAPS_USAGE);
    process.exit(help ? 0 : 2);
  }
  const result = await checkAgainstBase(base);
  for (const line of result.lines) (result.exitCode === 2 ? console.error : console.log)(line);
  process.exit(result.exitCode);
}
