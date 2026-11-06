// The fixture gate of `bun run fidelity`: a known-gap ratchet. The slides that fail today are
// listed in `FIXTURE_KNOWN_GAPS` (thresholds.ts), with each threshold they miss and its measured
// value. The gate fails when a slide fails in a way the list does not name, scores worse or
// better than its value by more than its margin, or passes now. A listed slide still reports
// "fail (known gap)", never a pass. `--update-gaps` rewrites the list from a run (`writeGaps`).

import { readFileSync, writeFileSync } from 'node:fs';
import {
  GAP_MARGINS,
  THRESHOLDS,
  type FailureKind,
  type KnownGap,
  type KnownGapList,
} from './thresholds.ts';
import type { DeckResult, SlideResult } from './report.ts';

/** `deck/slide`, for example `text-runs/1`. */
export const slideKey = (deck: string, slide: number): string => `${deck}/${slide}`;

const KINDS: readonly FailureKind[] = ['pixels', 'line-breaks', 'baseline', 'extra-text'];

/** The thresholds a compared slide misses, in a fixed order. Empty for a pass or a gap. */
export function failureKinds(s: SlideResult): FailureKind[] {
  if (s.status !== 'fail') return [];
  const kinds: FailureKind[] = [];
  if ((s.pixelMismatch ?? 0) > THRESHOLDS.maxPixelMismatch) kinds.push('pixels');
  if ((s.lineMatch ?? 1) < THRESHOLDS.minLineBreakMatch) kinds.push('line-breaks');
  if ((s.baselineDelta ?? 0) > THRESHOLDS.maxBaselineDeltaPt) kinds.push('baseline');
  if ((s.extraLines?.length ?? 0) > THRESHOLDS.maxExtraLines) kinds.push('extra-text');
  return kinds;
}

const round = (value: number, digits: number): number =>
  Math.round(value * 10 ** digits) / 10 ** digits;

/** The entry that records a failing slide's values, or undefined for a slide that passes. */
export function gapEntry(s: SlideResult): KnownGap | undefined {
  const kinds = failureKinds(s);
  if (kinds.length === 0) return undefined;
  const entry: { -readonly [K in keyof KnownGap]: KnownGap[K] } = {};
  if (kinds.includes('pixels')) entry.pixels = round(s.pixelMismatch ?? 0, 4);
  if (kinds.includes('line-breaks')) entry['line-breaks'] = round(s.lineMatch ?? 1, 4);
  if (kinds.includes('baseline')) entry.baseline = round(s.baselineDelta ?? 0, 2);
  if (kinds.includes('extra-text')) entry['extra-text'] = s.extraLines?.length ?? 0;
  if (s.linesTotal) entry.lines = s.linesTotal;
  return entry;
}

/** The kinds a known-gap entry lists. */
export const gapKinds = (gap: KnownGap): FailureKind[] => KINDS.filter((k) => gap[k] !== undefined);

const pct = (v: number): string => `${(v * 100).toFixed(2)} %`;

/** A slide's score for one kind, as the report writes it. */
function scoreText(s: SlideResult, kind: FailureKind): string {
  if (kind === 'pixels') return `pixels ${pct(s.pixelMismatch ?? 0)}`;
  if (kind === 'line-breaks')
    return `line breaks ${s.linesMatched ?? '?'} of ${s.linesTotal ?? '?'}`;
  if (kind === 'extra-text') return `${s.extraLines?.length ?? 0} extra lines`;
  return `baseline ${(s.baselineDelta ?? 0).toFixed(2)} pt`;
}

/** The threshold of one kind, as the report writes it. */
function thresholdText(kind: FailureKind): string {
  if (kind === 'pixels') return `≤ ${pct(THRESHOLDS.maxPixelMismatch)}`;
  if (kind === 'line-breaks') return 'all lines';
  if (kind === 'extra-text') return `≤ ${THRESHOLDS.maxExtraLines} extra lines`;
  return `≤ ${THRESHOLDS.maxBaselineDeltaPt} pt`;
}

/** A recorded value, as the report writes it. */
function recordText(kind: FailureKind, gap: KnownGap): string {
  if (kind === 'pixels') return `pixels ${pct(gap.pixels ?? 0)}`;
  if (kind === 'line-breaks') return `line breaks ${pct(gap['line-breaks'] ?? 0)}`;
  if (kind === 'extra-text') return `${gap['extra-text'] ?? 0} extra lines`;
  return `baseline ${(gap.baseline ?? 0).toFixed(2)} pt`;
}

/** How a measured value compares with its record: worse or better by more than the margin. */
function compareWithRecord(
  s: SlideResult,
  kind: FailureKind,
  gap: KnownGap
): 'worse' | 'better' | 'same' {
  const value =
    kind === 'pixels'
      ? s.pixelMismatch
      : kind === 'line-breaks'
        ? s.lineMatch
        : kind === 'extra-text'
          ? s.extraLines?.length
          : s.baselineDelta;
  const record = gap[kind];
  if (value === undefined || record === undefined) return 'same';
  // A higher value is worse, except for line breaks, where a higher share is better.
  const delta = kind === 'line-breaks' ? record - value : value - record;
  if (delta > GAP_MARGINS[kind] + 1e-12) return 'worse';
  if (delta < -GAP_MARGINS[kind] - 1e-12) return 'better';
  return 'same';
}

/** The per-slide report line of a failing slide: its scores, and its known gap if it has one. */
export function slideVerdict(s: SlideResult, gap: KnownGap | undefined): string {
  const kinds = failureKinds(s);
  if (kinds.length === 0) return '';
  const fails = kinds.map((k) => `${scoreText(s, k)} (threshold ${thresholdText(k)})`).join('; ');
  return gap
    ? `fails: ${fails}. Known gap: ${gapKinds(gap)
        .map((k) => recordText(k, gap))
        .join('; ')}`
    : `fails: ${fails}`;
}

/**
 * One problem of the gate. `record` problems go away with `--update-gaps`; `fix` ones do not.
 * `tools` says that the list was measured with other tools: record it in a change of the tools
 * alone (`check-known-gaps.ts`).
 */
export interface GateProblem {
  readonly kind: 'fix' | 'record' | 'tools';
  readonly text: string;
}

/** The tools that a known-gap list was measured with. */
export interface GapTools {
  readonly pdfjs: string;
  readonly chromium: string;
  readonly fonts: string;
}

const TOOL_KEYS = ['pdfjs', 'chromium', 'fonts'] as const;

/** The tools that differ between a list and a run, as `chromium a → b`. */
export function toolChanges(list: GapTools, tools: GapTools): string[] {
  return TOOL_KEYS.filter((k) => tools[k] !== list[k]).map(
    (k) => `${k} ${list[k] || 'none'} → ${tools[k]}`
  );
}

/** The reference lines that a deck checks over all its compared slides. */
export const deckLineCount = (deck: DeckResult): number =>
  deck.slides.reduce((n, s) => n + (s.linesTotal ?? 0), 0);

export interface GateResult {
  readonly problems: readonly GateProblem[];
  /** The slides that fail only in listed ways, at their recorded values. */
  readonly knownGaps: ReadonlySet<string>;
}

/**
 * Check the fixture results against the known gaps. `complete` is false when only some decks ran:
 * then a listed slide that did not run is no problem.
 */
export function checkFixtureGate(
  decks: readonly DeckResult[],
  list: KnownGapList,
  tools: GapTools,
  complete: boolean
): GateResult {
  const gaps = list.slides;
  const problems: GateProblem[] = [];
  const fix = (text: string): void => void problems.push({ kind: 'fix', text });
  const record = (text: string): void => void problems.push({ kind: 'record', text });
  const changed = toolChanges(list, tools);
  if (changed.length)
    problems.push({
      kind: 'tools',
      text: `the list was measured with other tools: ${changed.join(', ')}`,
    });
  const knownGaps = new Set<string>();
  const seen = new Set<string>();
  const unscored = new Set<string>();
  for (const deck of decks) {
    if (deck.foreignFonts?.length)
      fix(
        `${deck.name}: our PDF draws with fonts that the harness did not load: ${deck.foreignFonts.join(', ')}. A system font took part; give the glyphs an open face`
      );
    if (deck.foreignReferenceFonts?.length)
      fix(
        `${deck.name}: the reference PDF embeds fonts that are not open fonts of the deck: ${deck.foreignReferenceFonts.join(', ')}. Name open fonts in the deck, or list the deck in NON_OPEN_FONT_DECKS (local-fonts.ts)`
      );
    if (!deck.error && deck.reference === 'pdf') {
      const lines = deckLineCount(deck);
      const recorded = list.deckLines?.[deck.name];
      if (recorded === undefined)
        record(`${deck.name}: checks ${lines} lines, the list records none`);
      else if (lines < recorded)
        fix(
          `${deck.name}: checks ${lines} of its ${recorded} recorded lines: a change drops text from the check`
        );
      else if (lines > recorded)
        record(`${deck.name}: checks ${lines} lines, the list records ${recorded}`);
    }
    if (deck.error) {
      fix(`${deck.name}: the deck was not scored (${deck.error})`);
      unscored.add(deck.name);
    } else if (deck.reference !== 'pdf')
      fix(
        `${deck.name}: no usable reference (${(deck.referenceIssues ?? []).join('; ') || 'no PDF'}). Put its reference PDF next to it and record it in references.json`
      );
    for (const s of deck.slides) {
      const key = slideKey(deck.name, s.index);
      seen.add(key);
      if (s.selfConsistent === false) fix(`${key}: save and reopen paint another slide`);
      if (s.status !== 'pass' && s.status !== 'fail') {
        if (deck.reference === 'pdf' && !s.hidden) fix(`${key}: not compared (${s.cause})`);
        continue;
      }
      const kinds = failureKinds(s);
      if (s.status === 'fail' && kinds.length === 0) {
        fix(`${key}: fails, but no threshold names the failure`);
        continue;
      }
      const gap = Object.hasOwn(gaps, key) ? gaps[key] : undefined;
      const listed = gap ? gapKinds(gap) : [];
      const extra = kinds.filter((k) => !listed.includes(k));
      const gone = listed.filter((k) => !kinds.includes(k));
      const worse = kinds.filter(
        (k) => gap && listed.includes(k) && compareWithRecord(s, k, gap) === 'worse'
      );
      const better = kinds.filter(
        (k) => gap && listed.includes(k) && compareWithRecord(s, k, gap) === 'better'
      );
      for (const k of extra)
        fix(`${key}: fails on ${scoreText(s, k)} (threshold ${thresholdText(k)}), a new failure`);
      for (const k of worse)
        fix(`${key}: ${scoreText(s, k)} is worse than its known gap (${recordText(k, gap!)})`);
      const recount = gap !== undefined && (gap.lines ?? 0) !== (s.linesTotal ?? 0);
      if (gap && recount && (s.linesTotal ?? 0) < (gap.lines ?? 0))
        fix(
          `${key}: checks ${s.linesTotal ?? 0} of its ${gap.lines} recorded lines: text is missing`
        );
      else if (gap && recount)
        record(`${key}: checks ${s.linesTotal} lines, the list records ${gap.lines ?? 0}`);
      if (gone.length) record(`${key}: passes on ${gone.join(', ')} now`);
      if (better.length) record(`${key}: scores better than its known gap on ${better.join(', ')}`);
      if (
        gap &&
        kinds.length &&
        !extra.length &&
        !worse.length &&
        !better.length &&
        !gone.length &&
        !recount
      )
        knownGaps.add(key);
    }
  }
  if (complete) {
    for (const key of Object.keys(gaps))
      if (!seen.has(key) && !unscored.has(key.slice(0, key.lastIndexOf('/'))))
        record(`${key}: listed, but no such fixture slide exists`);
    const names = new Set(decks.map((d) => d.name));
    for (const deck of Object.keys(list.deckLines ?? {}))
      if (!names.has(deck)) record(`${deck}: a line count is listed, but no such fixture exists`);
  }
  return { problems, knownGaps };
}

/** Whether `head` allows a worse score than `base` for one kind. */
export function looser(kind: FailureKind, head: KnownGap, base: KnownGap): boolean {
  if (kind === 'pixels') return (head.pixels ?? 0) > (base.pixels ?? 0);
  if (kind === 'line-breaks') return (head['line-breaks'] ?? 1) < (base['line-breaks'] ?? 1);
  if (kind === 'extra-text') return (head['extra-text'] ?? 0) > (base['extra-text'] ?? 0);
  return (head.baseline ?? 0) > (base.baseline ?? 0);
}

/** How a new known-gap list differs from the old one, entry by entry. */
export interface ListChanges {
  /** The tools that changed, as `chromium a → b`. */
  readonly tools: readonly string[];
  /** Entries with a looser value or a new kind, and decks that check fewer lines. */
  readonly looser: readonly string[];
  /** New entries: slides that fail now and did not before. */
  readonly added: readonly string[];
  /** Entries that are gone: slides that pass now (or no longer exist). */
  readonly removed: readonly string[];
}

/** The looser, new, and removed entries of `next` against `old`. */
export function listChanges(old: KnownGapList, next: KnownGapList): ListChanges {
  const looserEntries: string[] = [];
  const added: string[] = [];
  for (const [key, gap] of Object.entries(next.slides)) {
    const was = Object.hasOwn(old.slides, key) ? old.slides[key] : undefined;
    if (!was) {
      added.push(gapEntryLine(key, gap));
      continue;
    }
    const worse = gapKinds(gap)
      .filter((k) => was[k] === undefined || looser(k, gap, was))
      .map((k) => `${k} ${was[k] ?? 'none'} → ${gap[k]}`);
    if ((gap.lines ?? 0) < (was.lines ?? 0)) worse.push(`lines ${was.lines} → ${gap.lines ?? 0}`);
    if (worse.length) looserEntries.push(`${key}: ${worse.join(', ')}`);
  }
  for (const [deck, lines] of Object.entries(next.deckLines)) {
    const was = old.deckLines?.[deck];
    if (was !== undefined && lines < was)
      looserEntries.push(`${deck}: checks fewer lines (${was} → ${lines})`);
  }
  const removed = [
    ...Object.keys(old.slides).filter((key) => !Object.hasOwn(next.slides, key)),
    ...Object.keys(old.deckLines ?? {})
      .filter((deck) => !Object.hasOwn(next.deckLines, deck))
      .map((deck) => `${deck} (its line count)`),
  ];
  return { tools: toolChanges(old, next), looser: looserEntries, added, removed };
}

/** What `--update-gaps` does with a run: the list it records, the changes, and whether to write. */
export interface GapUpdate {
  readonly next: KnownGapList;
  readonly changes: ListChanges;
  /** False when the run has new or looser entries and `acceptWorse` is not set. */
  readonly write: boolean;
}

/**
 * Plan `--update-gaps`: new entries and looser values are written only with `acceptWorse`.
 * `complete` is true when every fixture ran (`mergeRecorded`).
 */
export function planUpdate(
  list: KnownGapList,
  decks: readonly DeckResult[],
  tools: GapTools,
  acceptWorse: boolean,
  complete: boolean
): GapUpdate {
  const next = mergeRecorded(list, decks, tools, complete);
  const changes = listChanges(list, next);
  const worse = changes.looser.length > 0 || changes.added.length > 0;
  return { next, changes, write: acceptWorse || !worse };
}

/**
 * The list after a run: the entries of the decks that were scored come from the run. The entries
 * of decks that did not finish stay, and so do those of decks that did not run in a partial run.
 * A complete run drops the entries of fixtures that no longer exist. Only a complete run with no
 * deck error records new tools: a partial run measured too little to stamp them.
 */
export function mergeRecorded(
  list: KnownGapList,
  decks: readonly DeckResult[],
  tools: GapTools,
  complete: boolean
): KnownGapList {
  const scored = decks.filter((d) => !d.error);
  const ran = new Set(scored.map((d) => d.name));
  const failed = new Set(decks.filter((d) => d.error).map((d) => d.name));
  const keep = (deck: string): boolean => (complete ? failed.has(deck) : !ran.has(deck));
  const kept = Object.fromEntries(
    Object.entries(list.slides).filter(([key]) => keep(key.slice(0, key.lastIndexOf('/'))))
  );
  const keptLines = Object.fromEntries(
    Object.entries(list.deckLines ?? {}).filter(([deck]) => keep(deck))
  );
  const stamp = complete && !failed.size;
  const recorded = recordedList(
    scored,
    stamp ? tools : { pdfjs: list.pdfjs, chromium: list.chromium, fonts: list.fonts }
  );
  return {
    ...recorded,
    deckLines: { ...keptLines, ...recorded.deckLines },
    slides: { ...kept, ...recorded.slides },
  };
}

/** One known-gap entry as a line of `thresholds.ts`. */
export function gapEntryLine(key: string, gap: KnownGap): string {
  const parts = Object.entries(gap).map(([k, v]) => `${k.includes('-') ? `'${k}'` : k}: ${v}`);
  return `'${key}': { ${parts.join(', ')} },`;
}

/** The known-gap list that a run records: its tools and each failing slide. */
export function recordedList(decks: readonly DeckResult[], tools: GapTools): KnownGapList {
  const slides: Record<string, KnownGap> = {};
  const deckLines: Record<string, number> = {};
  for (const deck of decks) {
    if (!deck.error && deck.reference === 'pdf') deckLines[deck.name] = deckLineCount(deck);
    for (const s of deck.slides) {
      const gap = gapEntry(s);
      if (gap) slides[slideKey(deck.name, s.index)] = gap;
    }
  }
  return { ...tools, deckLines, slides };
}

/** The `FIXTURE_KNOWN_GAPS` constant as source text. */
export function listSource(list: KnownGapList): string {
  const entries = Object.keys(list.slides)
    .sort()
    .map((key) => `    ${gapEntryLine(key, list.slides[key])}`);
  const lines = Object.keys(list.deckLines)
    .sort()
    .map((deck) => `    '${deck}': ${list.deckLines[deck]},`);
  return [
    'export const FIXTURE_KNOWN_GAPS: KnownGapList = {',
    `  pdfjs: '${list.pdfjs}',`,
    `  chromium: '${list.chromium}',`,
    `  fonts: '${list.fonts}',`,
    lines.length ? `  deckLines: {\n${lines.join('\n')}\n  },` : '  deckLines: {},',
    entries.length ? `  slides: {\n${entries.join('\n')}\n  },` : '  slides: {},',
    '};',
  ].join('\n');
}

/** Replace `FIXTURE_KNOWN_GAPS` in the thresholds file with `list`, formatted as the repo is. */
export async function writeGaps(file: string, list: KnownGapList): Promise<void> {
  const text = readFileSync(file, 'utf8');
  const start = text.indexOf('export const FIXTURE_KNOWN_GAPS');
  const end = text.indexOf('\n};', start);
  if (start < 0 || end < 0) throw new Error(`${file} has no FIXTURE_KNOWN_GAPS constant`);
  const prettier = await import('prettier');
  const source = `${text.slice(0, start)}${listSource(list)}${text.slice(end + 3)}`;
  // The repository config, also for a file outside it (tests).
  const options = { ...(await prettier.resolveConfig(import.meta.path)), filepath: file };
  writeFileSync(file, await prettier.format(source, options));
}
