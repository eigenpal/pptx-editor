import { afterEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  checkFixtureGate as gateOf,
  deckLineCount,
  failureKinds,
  gapEntry,
  gapEntryLine,
  listChanges,
  listSource,
  mergeRecorded,
  planUpdate,
  recordedList,
  slideVerdict,
  writeGaps,
} from './known-gaps.ts';
import { FIXTURE_KNOWN_GAPS, type KnownGapList } from './thresholds.ts';
import { fixtureFontsIdentity } from './local-fonts.ts';
import { PDFJS_VERSION } from './pdf.ts';
import { pinnedChromium } from './harness-page.ts';
import type { DeckResult, SlideResult } from './report.ts';

const slide = (index: number, over: Partial<SlideResult> = {}): SlideResult => ({
  index,
  status: 'pass',
  pixelMismatch: 0,
  lineMatch: 1,
  linesMatched: 4,
  linesTotal: 4,
  baselineDelta: 0,
  selfConsistent: true,
  diagnostics: [],
  cause: '',
  ...over,
});

const deck = (name: string, slides: SlideResult[], over: Partial<DeckResult> = {}): DeckResult => ({
  name,
  reference: 'pdf',
  slides,
  ...over,
});

const tools = { pdfjs: '5', chromium: '1', fonts: 'f' };
const list = (
  slides: KnownGapList['slides'],
  deckLines: KnownGapList['deckLines'] = {}
): KnownGapList => ({ ...tools, deckLines, slides });
const baselineFail = { status: 'fail' as const, baselineDelta: 0.9 };
const texts = (r: ReturnType<typeof checkFixtureGate>): string[] => r.problems.map((p) => p.text);

/**
 * The gate, with each deck's checked lines recorded as the run counts them when the list records
 * none, so a test sees only the problems it is about.
 */
function checkFixtureGate(
  decks: readonly DeckResult[],
  l: KnownGapList,
  t: typeof tools,
  complete: boolean
): ReturnType<typeof gateOf> {
  const counted = Object.fromEntries(
    decks.filter((d) => !d.error && d.reference === 'pdf').map((d) => [d.name, deckLineCount(d)])
  );
  return gateOf(
    decks,
    Object.keys(l.deckLines).length ? l : { ...l, deckLines: counted },
    t,
    complete
  );
}

describe('failure kinds and entries', () => {
  test('name each threshold a failing slide misses, and record its values', () => {
    expect(failureKinds(slide(1))).toEqual([]);
    const s = slide(1, { status: 'fail', pixelMismatch: 0.02, lineMatch: 0.5, baselineDelta: 3 });
    expect(failureKinds(s)).toEqual(['pixels', 'line-breaks', 'baseline']);
    expect(gapEntry(s)).toEqual({ pixels: 0.02, 'line-breaks': 0.5, baseline: 3, lines: 4 });
    expect(gapEntry(slide(1))).toBeUndefined();
    expect(failureKinds(slide(1, { status: 'stale' }))).toEqual([]);
    expect(gapEntryLine('a/1', { pixels: 0.02, 'line-breaks': 0.5 })).toBe(
      "'a/1': { pixels: 0.02, 'line-breaks': 0.5 },"
    );
  });

  test('extra lines of our page are a failure kind with a count', () => {
    const s = slide(1, { status: 'fail', extraLines: ['Our own text', 'More text'] });
    expect(failureKinds(s)).toEqual(['extra-text']);
    expect(gapEntry(s)).toEqual({ 'extra-text': 2, lines: 4 });
    const gate = checkFixtureGate(
      [deck('a', [s])],
      list({ 'a/1': { 'extra-text': 1, lines: 4 } }),
      tools,
      true
    );
    expect(texts(gate)).toEqual(['a/1: 2 extra lines is worse than its known gap (1 extra lines)']);
  });

  test('a verdict names the scores and the known gap', () => {
    const s = slide(2, baselineFail);
    expect(slideVerdict(s, undefined)).toBe('fails: baseline 0.90 pt (threshold ≤ 0.5 pt)');
    expect(slideVerdict(s, { baseline: 0.9 })).toContain('Known gap: baseline 0.90 pt');
  });
});

describe('the fixture gate', () => {
  test('passes when every failure is listed at its value', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1), slide(2, baselineFail)])],
      list({ 'a/2': { baseline: 0.9, lines: 4 } }),
      tools,
      true
    );
    expect(r.problems).toEqual([]);
    expect([...r.knownGaps]).toEqual(['a/2']);
  });

  test('a new failure and a worse value need a fix; a better value or a pass needs a record', () => {
    const r = checkFixtureGate(
      [
        deck('a', [
          slide(1, { status: 'fail', pixelMismatch: 0.02 }),
          slide(2, { status: 'fail', baselineDelta: 1.5 }),
          slide(3, { status: 'fail', baselineDelta: 0.7 }),
          slide(4),
        ]),
      ],
      list({
        'a/2': { baseline: 0.9, lines: 4 },
        'a/3': { baseline: 0.9, lines: 4 },
        'a/4': { pixels: 0.01, lines: 4 },
      }),
      tools,
      true
    );
    expect(r.problems.map((p) => p.kind)).toEqual(['fix', 'fix', 'record', 'record']);
    expect(texts(r)[0]).toContain('a/1: fails on pixels 2.00 %');
    expect(texts(r)[1]).toContain('a/2: baseline 1.50 pt is worse than its known gap');
    expect(texts(r)[2]).toContain('a/3: scores better than its known gap on baseline');
    expect(texts(r)[3]).toContain('a/4: passes on pixels now');
  });

  test('fewer lines than recorded need a fix: text that disappears is no improvement', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1, { ...baselineFail, linesTotal: 2 })])],
      list({ 'a/1': { baseline: 0.9, lines: 4 } }),
      tools,
      true
    );
    expect(r.problems).toEqual([
      { kind: 'fix', text: 'a/1: checks 2 of its 4 recorded lines: text is missing' },
    ]);
  });

  test('references, open errors, save consistency, and other tools are problems', () => {
    const r = checkFixtureGate(
      [
        deck('a', [slide(1, { selfConsistent: false })]),
        deck('b', [], { error: 'boom' }),
        deck('c', [slide(1, { status: 'no-reference', cause: 'no PDF' })], {
          reference: 'none',
          referenceIssues: ['no reference PDF'],
        }),
      ],
      list({ 'b/1': { pixels: 0.1 }, 'z/1': { pixels: 0.1 } }),
      { ...tools, pdfjs: '6' },
      true
    );
    expect(texts(r)).toEqual([
      'the list was measured with other tools: pdfjs 5 → 6',
      'a/1: save and reopen paint another slide',
      'b: the deck was not scored (boom)',
      'c: no usable reference (no reference PDF). Put its reference PDF next to it and record it in references.json',
      'z/1: listed, but no such fixture slide exists',
    ]);
  });

  test('a deck that checks fewer lines than recorded needs a fix; more or none needs a record', () => {
    const decks = [deck('a', [slide(1)]), deck('b', [slide(1)]), deck('c', [slide(1)])];
    const r = gateOf(decks, list({}, { a: 6, b: 2 }), tools, true);
    expect(r.problems).toEqual([
      {
        kind: 'fix',
        text: 'a: checks 4 of its 6 recorded lines: a change drops text from the check',
      },
      { kind: 'record', text: 'b: checks 4 lines, the list records 2' },
      { kind: 'record', text: 'c: checks 4 lines, the list records none' },
    ]);
    expect(
      gateOf(decks, list({}, { a: 4, b: 4, c: 4 }), { ...tools, chromium: '2' }, true).problems
    ).toEqual([{ kind: 'tools', text: 'the list was measured with other tools: chromium 1 → 2' }]);
  });

  test('a failed slide that no threshold names needs a fix', () => {
    const r = checkFixtureGate([deck('a', [slide(1, { status: 'fail' })])], list({}), tools, true);
    expect(r.problems).toEqual([
      { kind: 'fix', text: 'a/1: fails, but no threshold names the failure' },
    ]);
  });

  test('fonts that the harness did not load, and reference fonts that are not open, need a fix', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1)], { foreignFonts: ['Helvetica'], foreignReferenceFonts: ['ArialMT'] })],
      list({}),
      tools,
      true
    );
    expect(r.problems.map((p) => p.kind)).toEqual(['fix', 'fix']);
    expect(texts(r)[0]).toStartWith(
      'a: our PDF draws with fonts that the harness did not load: Helvetica.'
    );
    expect(texts(r)[1]).toStartWith(
      'a: the reference PDF embeds fonts that are not open fonts of the deck: ArialMT.'
    );
  });

  test('a partial run does not ask for the slides it did not score', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1)])],
      list({ 'z/1': { pixels: 0.1 } }),
      tools,
      false
    );
    expect(r.problems).toEqual([]);
  });
});

describe('--update-gaps', () => {
  let dir: string | undefined;
  afterEach(() => dir && rmSync(dir, { recursive: true, force: true }));

  test('writes the list that a run records, formatted, and reads back the same', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pptx-gaps-'));
    const file = join(dir, 'thresholds.ts');
    writeFileSync(
      file,
      `export interface KnownGapList { slides: object }\n\n${listSource(list({}))}\n\nexport const AFTER = 1;\n`
    );
    const recorded = recordedList([deck('a', [slide(1), slide(2, baselineFail)])], tools);
    await writeGaps(file, recorded);
    const text = readFileSync(file, 'utf8');
    expect(text).toContain("'a/2': { baseline: 0.9, lines: 4 },");
    expect(text).toContain('export const AFTER = 1;');
    const mod = (await import(file)) as { FIXTURE_KNOWN_GAPS: KnownGapList };
    expect(mod.FIXTURE_KNOWN_GAPS).toEqual(recorded);
  });

  test('a partial run, or a deck that did not finish, keeps the entries of the decks it lacks', () => {
    const old = list({
      'a/1': { pixels: 0.1 },
      'b/1': { pixels: 0.2 },
      'c/1': { pixels: 0.3 },
    });
    const next = mergeRecorded(
      old,
      [
        deck('a', [slide(1)]),
        deck('b', [], { error: 'the browser crashed while rendering the deck' }),
      ],
      tools
    );
    expect(next.slides).toEqual({ 'b/1': { pixels: 0.2 }, 'c/1': { pixels: 0.3 } });
  });

  test('listChanges names looser, new, and removed entries', () => {
    const old = list({ 'a/1': { pixels: 0.1, lines: 4 }, 'b/1': { pixels: 0.2 } });
    const next = list({ 'a/1': { pixels: 0.2, baseline: 1, lines: 3 }, 'c/1': { pixels: 0.3 } });
    expect(listChanges(old, next)).toEqual({
      tools: [],
      looser: ['a/1: pixels 0.1 → 0.2, baseline none → 1, lines 4 → 3'],
      added: ["'c/1': { pixels: 0.3 },"],
      removed: ['b/1'],
    });
    expect(listChanges(old, list({ 'a/1': { pixels: 0.05, lines: 4 } }))).toEqual({
      tools: [],
      looser: [],
      added: [],
      removed: ['b/1'],
    });
  });

  test('listChanges names a change of the tools and fewer lines in a deck', () => {
    const old = list({}, { a: 10 });
    const next = { ...list({}, { a: 8 }), chromium: '2' };
    expect(listChanges(old, next)).toMatchObject({
      tools: ['chromium 1 → 2'],
      looser: ['a: checks fewer lines (10 → 8)'],
    });
  });

  test('--update-gaps writes better scores, and worse ones only with --accept-worse', () => {
    const old = list({ 'a/1': { baseline: 0.9, lines: 4 } }, { a: 4 });
    const better = [deck('a', [slide(1)])];
    expect(planUpdate(old, better, tools, false)).toMatchObject({
      write: true,
      changes: { removed: ['a/1'] },
    });
    const worse = [deck('a', [slide(1, { status: 'fail', baselineDelta: 2 })])];
    expect(planUpdate(old, worse, tools, false).write).toBe(false);
    expect(planUpdate(old, worse, tools, true).write).toBe(true);
    const added = [deck('a', [slide(1)]), deck('b', [slide(1, baselineFail)])];
    expect(planUpdate(old, added, tools, false)).toMatchObject({
      write: false,
      changes: { added: ["'b/1': { baseline: 0.9, lines: 4 },"] },
    });
  });

  test('the committed list was measured with the tools and fonts in use', () => {
    expect(FIXTURE_KNOWN_GAPS.pdfjs).toBe(PDFJS_VERSION);
    expect(FIXTURE_KNOWN_GAPS.fonts).toBe(fixtureFontsIdentity());
    expect(FIXTURE_KNOWN_GAPS.chromium).toBe(pinnedChromium());
  });
});
