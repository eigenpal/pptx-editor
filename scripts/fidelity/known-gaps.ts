// The fixture gate of `bun run fidelity`: a known-gap ratchet. Every fixture slide is compared with its
// reference PDF at the thresholds in `thresholds.ts`. The slides that fail today are listed in
// `FIXTURE_KNOWN_GAPS`, per environment, with each threshold they miss and its measured value.
// The list is a ratchet: the gate fails when a slide fails in a way the list does not name, when
// a listed slide scores worse or better than its value by more than its margin, and when a
// listed failure is gone, until the same change records it. A listed slide still reports "fail
// (known gap)", never a pass. `check-known-gaps.ts` refuses, in CI, a change that adds an entry
// or loosens a value without a new reference.

import {
  GAP_MARGINS,
  THRESHOLDS,
  type FailureKind,
  type FidelityEnvironment,
  type KnownGap,
} from './thresholds.ts';
import type { DeckResult, SlideResult } from './report.ts';
import { sizeLimitText } from './compare.ts';

/** `deck/slide`, for example `text-runs/1`. */
export const slideKey = (deck: string, slide: number): string => `${deck}/${slide}`;

/**
 * The environment whose known-gap list applies. Scores depend on the fonts and the rasterizer:
 * CI is Linux without the reference fonts; `macos-reference` is macOS with the reference fonts.
 * Other environments have no list, and their fidelity results are reported, not gated.
 */
export function fidelityEnvironment(
  platform: string,
  referenceFonts: boolean
): FidelityEnvironment | undefined {
  if (platform === 'linux' && !referenceFonts) return 'linux';
  if (platform === 'darwin' && referenceFonts) return 'macos-reference';
  return undefined;
}

/** The thresholds a compared slide misses, in a fixed order. Empty for a pass or a gap. */
export function failureKinds(s: SlideResult): FailureKind[] {
  if (s.status !== 'fail') return [];
  if (s.sizeMismatch) return ['size'];
  const kinds: FailureKind[] = [];
  if ((s.pixelMismatch ?? 0) > THRESHOLDS.maxPixelMismatch) kinds.push('pixels');
  if ((s.lineMatch ?? 1) < THRESHOLDS.minLineBreakMatch) kinds.push('line-breaks');
  if ((s.baselineDelta ?? 0) > THRESHOLDS.maxBaselineDeltaPt) kinds.push('baseline');
  return kinds;
}

const round = (value: number, digits: number): number =>
  Math.round(value * 10 ** digits) / 10 ** digits;

/**
 * The known-gap entry that records a failing slide's measured values, or undefined for a slide
 * that passes. Pixels and line breaks keep 4 decimals, baselines 2. `lines` is the count of
 * checked lines of every slide that has one, also when only its pixels fail: so an entry that
 * drops its line kinds keeps its line count, and text that disappears cannot pass as a fix.
 */
export function gapEntry(s: SlideResult): KnownGap | undefined {
  const kinds = failureKinds(s);
  if (kinds.length === 0) return undefined;
  if (kinds.includes('size')) return { size: true };
  const entry: { -readonly [K in keyof KnownGap]: KnownGap[K] } = {};
  if (kinds.includes('pixels')) entry.pixels = round(s.pixelMismatch ?? 0, 4);
  if (kinds.includes('line-breaks')) entry['line-breaks'] = round(s.lineMatch ?? 1, 4);
  if (kinds.includes('baseline')) entry.baseline = round(s.baselineDelta ?? 0, 2);
  if (s.linesTotal) entry.lines = s.linesTotal;
  return entry;
}

/** The kinds a known-gap entry lists, in the order of `failureKinds`. */
export const gapKinds = (gap: KnownGap): FailureKind[] =>
  (['size', 'pixels', 'line-breaks', 'baseline'] as const).filter((k) => gap[k] !== undefined);

const pct = (v: number): string => `${(v * 100).toFixed(2)} %`;

/** A slide's score for one kind, as the report writes it. */
function scoreText(s: SlideResult, kind: FailureKind): string {
  // The cause says which image is wrong, and the detail gives the sizes (compare.ts).
  if (kind === 'size') {
    // The size cause is the first cause; the causes after it are the slide's other causes.
    const first = s.cause.split('; ')[0];
    const cause = first.startsWith('slide size differs') ? first : 'slide size differs';
    return s.sizeDetail ? `${cause} — ${s.sizeDetail}` : cause;
  }
  if (kind === 'pixels') return `pixels ${pct(s.pixelMismatch ?? 0)}`;
  if (kind === 'line-breaks')
    return `line breaks ${s.linesMatched ?? '?'} of ${s.linesTotal ?? '?'} (${pct(s.lineMatch ?? 1)})`;
  return `baseline ${(s.baselineDelta ?? 0).toFixed(2)} pt`;
}

/** The threshold of one kind, as the report writes it. */
function thresholdText(kind: FailureKind): string {
  if (kind === 'size') return sizeLimitText();
  if (kind === 'pixels') return `≤ ${pct(THRESHOLDS.maxPixelMismatch)}`;
  if (kind === 'line-breaks') return `≥ ${pct(THRESHOLDS.minLineBreakMatch)}`;
  return `≤ ${THRESHOLDS.maxBaselineDeltaPt} pt`;
}

/** A recorded value, as the report writes it. */
function recordText(kind: FailureKind, gap: KnownGap): string {
  if (kind === 'size') return 'size';
  if (kind === 'pixels') return `pixels ${pct(gap.pixels ?? 0)}`;
  if (kind === 'line-breaks') return `line breaks ${pct(gap['line-breaks'] ?? 0)}`;
  return `baseline ${(gap.baseline ?? 0).toFixed(2)} pt`;
}

/** How a measured value compares with its record: worse or better by more than the margin. */
function compareWithRecord(
  s: SlideResult,
  kind: FailureKind,
  gap: KnownGap
): 'worse' | 'better' | 'same' {
  if (kind === 'size') return 'same';
  const value =
    kind === 'pixels' ? s.pixelMismatch : kind === 'line-breaks' ? s.lineMatch : s.baselineDelta;
  const record = gap[kind];
  if (value === undefined || record === undefined) return 'same';
  // A higher value is worse, except for line breaks, where a higher share is better.
  const delta = kind === 'line-breaks' ? record - value : value - record;
  if (delta > GAP_MARGINS[kind] + 1e-12) return 'worse';
  if (delta < -GAP_MARGINS[kind] - 1e-12) return 'better';
  return 'same';
}

/** Whether a listed line kind can still be measured on the slide. */
function measured(s: SlideResult, kind: FailureKind): boolean {
  if (kind === 'line-breaks') return s.lineMatch !== undefined;
  if (kind === 'baseline') return s.baselineDelta !== undefined;
  return true;
}

/** One known-gap entry as a line of `thresholds.ts`. */
export function gapEntryLine(key: string, gap: KnownGap): string {
  const parts = Object.entries(gap).map(([k, v]) => `${k.includes('-') ? `'${k}'` : k}: ${v}`);
  return `'${key}': { ${parts.join(', ')} },`;
}

/** The per-slide report line of a failing slide: its scores, and its known gap if it has one. */
export function slideVerdict(s: SlideResult, gap: KnownGap | undefined): string {
  const kinds = failureKinds(s);
  if (kinds.length === 0) return '';
  const fails = kinds
    .map(
      (k) => `${scoreText(s, k)} (${k === 'size' ? 'size check' : 'threshold'} ${thresholdText(k)})`
    )
    .join('; ');
  return gap
    ? `fails: ${fails}. Known gap: ${gapKinds(gap)
        .map((k) => recordText(k, gap))
        .join('; ')}`
    : `fails: ${fails}`;
}

export interface GateOptions {
  /** Every fixture ran (no `--deck`), so a listed slide that does not exist is a problem. */
  readonly complete: boolean;
  /** The list's name in messages, for example `FIXTURE_KNOWN_GAPS.linux`. */
  readonly list?: string;
  /**
   * The tools of this run that differ from the ones the list was measured with, for example
   * `Poppler 26.11.0 (the list: 26.10.0)`. A changed score may come from them.
   */
  readonly toolChanges?: readonly string[];
}

export interface GateResult {
  /** Why the gate fails. Empty when it passes. */
  readonly problems: readonly string[];
  /** Facts that do not fail the gate, for example a tool change. */
  readonly notes: readonly string[];
  /** The slides that fail only in listed ways, at their recorded values. */
  readonly knownGaps: ReadonlySet<string>;
}

/** The gate problem for a deck without a usable reference, with its reasons and the next step. */
function referenceProblem(deck: DeckResult): string {
  const reasons = deck.referenceIssues?.length ? deck.referenceIssues.join('; ') : 'no PDF';
  return `${deck.name}: no usable reference (${reasons}). It needs a reference PDF next to it and its entry in references.json.`;
}

/**
 * Check the fixture results against the known gaps. Without a list (`gaps` undefined), slide
 * failures are reported but do not fail the gate; references, open errors, and save consistency
 * still do.
 */
export function checkFixtureGate(
  decks: readonly DeckResult[],
  gaps: Readonly<Record<string, KnownGap>> | undefined,
  options: GateOptions
): GateResult {
  const list = options.list ?? 'FIXTURE_KNOWN_GAPS';
  const tools = options.toolChanges?.length
    ? ` This run uses other tools than the list: ${options.toolChanges.join(', ')}. If only the tools changed, record the list again for them.`
    : '';
  const problems: string[] = [];
  const notes: string[] = options.toolChanges?.length
    ? [
        `this run uses other tools than ${list} was measured with: ${options.toolChanges.join(', ')}`,
      ]
    : [];
  const knownGaps = new Set<string>();
  const seen = new Set<string>();
  const unopened = new Set<string>();
  for (const deck of decks) {
    if (deck.error) {
      problems.push(`${deck.name}: the harness could not score the deck (${deck.error})`);
      unopened.add(deck.name);
    } else if (deck.reference !== 'pdf') problems.push(referenceProblem(deck));
    for (const s of deck.slides) {
      const key = slideKey(deck.name, s.index);
      seen.add(key);
      if (s.selfConsistent === false)
        problems.push(`${key}: save → reopen → render differs from the first render`);
      if (
        deck.reference === 'pdf' &&
        !s.hidden &&
        (s.status === 'no-reference' || s.status === 'stale')
      )
        problems.push(`${key}: not compared (${s.cause || 'no PDF page'})`);
      if (s.status === 'fail' && failureKinds(s).length === 0)
        problems.push(`${key}: fails, but on no threshold that the gate knows`);
      if (!gaps || (s.status !== 'pass' && s.status !== 'fail')) continue;
      const kinds = failureKinds(s);
      const gap = Object.hasOwn(gaps, key) ? gaps[key] : undefined;
      const listed = gap ? gapKinds(gap) : [];
      // A size entry gates nothing else. When the size is fixed, the scores that replace it are
      // no new failures: the slide passes on size now, and the change records its new entry.
      const sizeFixed = gap?.size === true && !kinds.includes('size');
      const extra = sizeFixed ? [] : kinds.filter((k) => !listed.includes(k));
      // A size mismatch gates nothing else, so the other listed kinds are unknown, not gone.
      const sized = kinds.includes('size');
      // A listed line kind that the slide no longer measures (its text is gone, or no line
      // matches) did not improve: it is a problem of its own, never "passes now".
      const lost = sized ? [] : listed.filter((k) => !kinds.includes(k) && !measured(s, k));
      const gone = sized ? [] : listed.filter((k) => !kinds.includes(k) && measured(s, k));
      const worse = gap
        ? kinds.filter((k) => listed.includes(k) && compareWithRecord(s, k, gap) === 'worse')
        : [];
      const better = gap
        ? kinds.filter((k) => listed.includes(k) && compareWithRecord(s, k, gap) === 'better')
        : [];
      const next = gapEntry(s);
      for (const k of extra)
        problems.push(
          k === 'size'
            ? // The slide content cannot change a size: the render or the raster is wrong.
              `${key}: ${scoreText(s, k)} (size check ${thresholdText(k)}), which is not a known gap. Find which size is wrong: the render size comes from renderSizePx and examples/vite/src/harness.ts, the raster size from expectedRasterSizePx (CONTRIBUTING.md, "Fidelity harness"). Only a new fixture, a changed deck, or a changed environment may add ${gapEntryLine(key, next ?? {})}${tools}`
            : `${key}: fails on ${scoreText(s, k)} (threshold ${thresholdText(k)}), which is not a known gap. Fix the slide; only a new fixture, a changed deck, or a changed environment may add ${gapEntryLine(key, next ?? {})}${tools}`
        );
      // Fewer checked lines than recorded: part of the text is gone, so the line scores of the
      // rest cannot count as better.
      const fewer =
        gap?.lines !== undefined && !sized && (s.linesTotal ?? 0) < gap.lines
          ? [
              `${key}: checks ${s.linesTotal ?? 0} of its ${gap.lines} recorded lines now. Find where the text went; a known gap does not end this way. Only if this change edits the deck (or exports it with another font set), record it: in ${list}, ${next ? `set ${gapEntryLine(key, next)}` : `remove '${key}'`}`,
            ]
          : [];
      problems.push(...fewer);
      // Another line count is a new record, so that a later loss of those lines shows. A pixels
      // entry of a slide with text records its count too (`gapEntry`).
      const recount =
        gap !== undefined && !sized && (s.linesTotal ?? 0) > 0 && gap.lines !== s.linesTotal;
      for (const k of lost)
        problems.push(
          `${key}: its ${k === 'baseline' ? 'baseline is' : 'line breaks are'} no longer measured: the slide has no ${k === 'baseline' ? 'matched' : 'checked'} line now (${s.linesMatched ?? 0} of ${s.linesTotal ?? 0} lines). Find where the text went; a known gap does not end this way`
        );
      for (const k of worse)
        problems.push(
          `${key}: ${scoreText(s, k)} is worse than its known gap (${recordText(k, gap!)}). Fix the regression: a known gap does not loosen.${tools}`
        );
      const quiet = !extra.length && !worse.length && !lost.length && !fewer.length;
      if ((gone.length || better.length) && quiet)
        problems.push(
          `${key}: ${gone.length ? `passes on ${gone.join(', ')} now` : `scores better than its known gap on ${better.join(', ')}`}. Record it in this change: in ${list}, ${next ? `set ${gapEntryLine(key, next)}` : `remove '${key}'`}${tools}`
        );
      else if (recount && quiet && next)
        problems.push(
          `${key}: checks ${s.linesTotal} lines, and its entry records ${gap.lines === undefined ? 'no line count' : gap.lines}. Record it in this change: in ${list}, set ${gapEntryLine(key, next)}${tools}`
        );
      if (
        gap &&
        kinds.length &&
        !extra.length &&
        !worse.length &&
        !better.length &&
        !lost.length &&
        !fewer.length &&
        !recount
      )
        knownGaps.add(key);
    }
  }
  if (gaps && options.complete)
    for (const key of Object.keys(gaps))
      if (!seen.has(key) && !unopened.has(key.slice(0, key.lastIndexOf('/'))))
        problems.push(`${key}: listed in ${list}, but no such fixture slide exists`);
  return { problems, notes, knownGaps };
}

/** The whole known-gap list that this run's results record, as lines of `thresholds.ts`. */
export function recordedGaps(decks: readonly DeckResult[]): string[] {
  const lines: string[] = [];
  for (const deck of decks)
    for (const s of deck.slides) {
      const gap = gapEntry(s);
      if (gap) lines.push(gapEntryLine(slideKey(deck.name, s.index), gap));
    }
  return lines;
}

/** How this run's tools differ from the ones a list was measured with, for the gate's notes. */
export function listToolChanges(
  recorded: { readonly poppler: string; readonly chromium: string } | undefined,
  poppler: string | null,
  chromium: string | null
): string[] {
  if (!recorded) return [];
  return [
    poppler !== recorded.poppler
      ? `Poppler ${poppler ?? 'none'} (the list: ${recorded.poppler})`
      : '',
    chromium !== recorded.chromium
      ? `Chromium ${chromium ?? 'unknown'} (the list: ${recorded.chromium})`
      : '',
  ].filter(Boolean);
}

/** The exit code of the fixture gate: 1 with problems, 3 when no slide score was gated, else 0. */
export function gateExitCode(problems: number, gated: boolean): 0 | 1 | 3 {
  if (problems) return 1;
  return gated ? 0 : 3;
}
