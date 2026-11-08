// The fixture gate of `bun run fidelity`: a known-gap ratchet. `FIXTURE_KNOWN_GAPS`
// (thresholds.ts) lists, for each slide that fails today, the checks it fails: never their
// values, because rasters differ slightly between machines. The gate fails when a slide fails a
// check that its entry does not list, or passes a listed check now (remove it). A listed slide
// still reports "fail (known gap)", never a pass. `--update-gaps` rewrites the list (`writeGaps`).

import { readFileSync, writeFileSync } from 'node:fs';
import { THRESHOLDS, type FailureKind, type KnownGap, type KnownGapList } from './thresholds.ts';
import type { DeckResult, SlideResult } from './report.ts';

/** `deck/slide`, for example `text-runs/1`. */
export const slideKey = (deck: string, slide: number): string => `${deck}/${slide}`;

const KINDS: readonly FailureKind[] = [
  'pixels',
  'line-breaks',
  'baseline',
  'line-position',
  'extra-text',
];

/** The thresholds a compared slide misses, in a fixed order. Empty for a pass or a gap. */
export function failureKinds(s: SlideResult): FailureKind[] {
  if (s.status !== 'fail') return [];
  const kinds: FailureKind[] = [];
  if ((s.pixelMismatch ?? 0) > THRESHOLDS.maxPixelMismatch) kinds.push('pixels');
  if ((s.lineMatch ?? 1) < THRESHOLDS.minLineBreakMatch) kinds.push('line-breaks');
  if ((s.baselineDelta ?? 0) > THRESHOLDS.maxBaselineDeltaPt) kinds.push('baseline');
  if ((s.lineShift ?? 0) > THRESHOLDS.maxLineShiftPt) kinds.push('line-position');
  if ((s.extraLines?.length ?? 0) > THRESHOLDS.maxExtraLines) kinds.push('extra-text');
  return kinds;
}

/** The entry of a failing slide: the checks it fails. Undefined for a slide that passes. */
export function gapEntry(s: SlideResult): KnownGap | undefined {
  const kinds = failureKinds(s);
  return kinds.length ? kinds : undefined;
}

/** The kinds a known-gap entry lists. */
export const gapKinds = (gap: KnownGap): FailureKind[] => KINDS.filter((k) => gap.includes(k));

const pct = (v: number): string => `${(v * 100).toFixed(2)} %`;

/** A slide's score for one kind, as the report writes it. */
function scoreText(s: SlideResult, kind: FailureKind): string {
  if (kind === 'pixels') return `pixels ${pct(s.pixelMismatch ?? 0)}`;
  if (kind === 'line-breaks')
    return `line breaks ${s.linesMatched ?? '?'} of ${s.linesTotal ?? '?'}`;
  if (kind === 'extra-text') return `${s.extraLines?.length ?? 0} extra lines`;
  if (kind === 'line-position') return `line position ${(s.lineShift ?? 0).toFixed(2)} pt`;
  return `baseline ${(s.baselineDelta ?? 0).toFixed(2)} pt`;
}

/** The threshold of one kind, as the report writes it. */
function thresholdText(kind: FailureKind): string {
  if (kind === 'pixels') return `≤ ${pct(THRESHOLDS.maxPixelMismatch)}`;
  if (kind === 'line-breaks') return 'all lines';
  if (kind === 'extra-text') return `≤ ${THRESHOLDS.maxExtraLines} extra lines`;
  if (kind === 'line-position') return `≤ ${THRESHOLDS.maxLineShiftPt} pt`;
  return `≤ ${THRESHOLDS.maxBaselineDeltaPt} pt`;
}

/** The per-slide report line of a failing slide: its scores, and its known gap if it has one. */
export function slideVerdict(s: SlideResult, gap: KnownGap | undefined): string {
  const kinds = failureKinds(s);
  if (kinds.length === 0) return '';
  const fails = kinds.map((k) => `${scoreText(s, k)} (threshold ${thresholdText(k)})`).join('; ');
  return gap ? `fails: ${fails}. Known gap: ${gapKinds(gap).join(', ')}` : `fails: ${fails}`;
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
  /** The slides that fail exactly the checks that their entries list. */
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
        `${deck.name}: the reference PDF embeds fonts that are not open fonts of the deck: ${deck.foreignReferenceFonts.join(', ')}. Name open fonts in the deck and export its reference PDF again`
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
      for (const k of extra)
        fix(`${key}: fails on ${scoreText(s, k)} (threshold ${thresholdText(k)}), a new failure`);
      if (gone.length) record(`${key}: passes on ${gone.join(', ')} now`);
      if (gap && kinds.length && !extra.length && !gone.length) knownGaps.add(key);
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

/** How a new known-gap list differs from the old one, entry by entry. */
export interface ListChanges {
  /** The tools that changed, as `chromium a → b`. */
  readonly tools: readonly string[];
  /** Entries with a new kind, and decks that check fewer lines. */
  readonly looser: readonly string[];
  /** New entries: slides that fail now and did not before. */
  readonly added: readonly string[];
  /** Entries that are gone: slides that pass now (or no longer exist). */
  readonly removed: readonly string[];
}

/** The new kinds, new entries, and removed entries of `next` against `old`. */
export function listChanges(old: KnownGapList, next: KnownGapList): ListChanges {
  const looserEntries: string[] = [];
  const added: string[] = [];
  for (const [key, gap] of Object.entries(next.slides)) {
    const was = Object.hasOwn(old.slides, key) ? old.slides[key] : undefined;
    if (!was) {
      added.push(gapEntryLine(key, gap));
      continue;
    }
    const more = gap.filter((k) => !was.includes(k));
    if (more.length) looserEntries.push(`${key}: fails ${more.join(', ')} now`);
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
  /** False when the run adds an entry or a kind and `acceptWorse` is not set. */
  readonly write: boolean;
}

/**
 * Plan `--update-gaps`: an added entry or kind is written only with `acceptWorse`.
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
  return `'${key}': [${gapKinds(gap)
    .map((k) => `'${k}'`)
    .join(', ')}],`;
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

/** The known-gap lists of `thresholds.ts`, by the constant's name. */
export type GapListName = 'FIXTURE_KNOWN_GAPS' | 'FIXTURE_EXPORT_KNOWN_GAPS';

/** A known-gap constant (`FIXTURE_KNOWN_GAPS` by default) as source text. */
export function listSource(list: KnownGapList, name: GapListName = 'FIXTURE_KNOWN_GAPS'): string {
  const entries = Object.keys(list.slides)
    .sort()
    .map((key) => `    ${gapEntryLine(key, list.slides[key])}`);
  const lines = Object.keys(list.deckLines)
    .sort()
    .map((deck) => `    '${deck}': ${list.deckLines[deck]},`);
  return [
    `export const ${name}: KnownGapList = {`,
    `  pdfjs: '${list.pdfjs}',`,
    `  chromium: '${list.chromium}',`,
    `  fonts: '${list.fonts}',`,
    lines.length ? `  deckLines: {\n${lines.join('\n')}\n  },` : '  deckLines: {},',
    entries.length ? `  slides: {\n${entries.join('\n')}\n  },` : '  slides: {},',
    '};',
  ].join('\n');
}

/**
 * Replace a known-gap constant (`FIXTURE_KNOWN_GAPS` by default) in the thresholds file with
 * `list`, formatted as the repo is.
 */
export async function writeGaps(
  file: string,
  list: KnownGapList,
  name: GapListName = 'FIXTURE_KNOWN_GAPS'
): Promise<void> {
  const text = readFileSync(file, 'utf8');
  const start = text.indexOf(`export const ${name}:`);
  const end = text.indexOf('\n};', start);
  if (start < 0 || end < 0) throw new Error(`${file} has no ${name} constant`);
  const prettier = await import('prettier');
  const source = `${text.slice(0, start)}${listSource(list, name)}${text.slice(end + 3)}`;
  // The repository config, also for a file outside it (tests).
  const options = { ...(await prettier.resolveConfig(import.meta.path)), filepath: file };
  writeFileSync(file, await prettier.format(source, options));
}
