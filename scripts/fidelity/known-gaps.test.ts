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
import { FIXTURE_EXPORT_KNOWN_GAPS, FIXTURE_KNOWN_GAPS, type KnownGapList } from './thresholds.ts';
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
  test('name each threshold a failing slide misses, and record only the kinds', () => {
    expect(failureKinds(slide(1))).toEqual([]);
    const s = slide(1, {
      status: 'fail',
      pixelMismatch: 0.02,
      lineMatch: 0.5,
      baselineDelta: 3,
      lineShift: 0.8,
    });
    expect(failureKinds(s)).toEqual(['pixels', 'line-breaks', 'baseline', 'line-position']);
    expect(gapEntry(s)).toEqual(['pixels', 'line-breaks', 'baseline', 'line-position']);
    expect(gapEntry(slide(1))).toBeUndefined();
    expect(failureKinds(slide(1, { status: 'stale' }))).toEqual([]);
    expect(gapEntryLine('a/1', ['pixels', 'line-breaks'])).toBe(
      "'a/1': ['pixels', 'line-breaks'],"
    );
  });

  test('extra lines of our page are a failure kind', () => {
    const s = slide(1, { status: 'fail', extraLines: ['Our own text', 'More text'] });
    expect(failureKinds(s)).toEqual(['extra-text']);
    const gate = checkFixtureGate([deck('a', [s])], list({}), tools, true);
    expect(texts(gate)).toEqual([
      'a/1: fails on 2 extra lines (threshold ≤ 0 extra lines), a new failure',
    ]);
  });

  test('a verdict names the scores and the known gap', () => {
    const s = slide(2, baselineFail);
    expect(slideVerdict(s, undefined)).toBe('fails: baseline 0.90 pt (threshold ≤ 0.5 pt)');
    expect(slideVerdict(s, ['baseline'])).toContain('Known gap: baseline');
  });
});

describe('the fixture gate', () => {
  test('passes when every failing check is listed, whatever its value', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1), slide(2, { status: 'fail', baselineDelta: 3 })])],
      list({ 'a/2': ['baseline'] }),
      tools,
      true
    );
    expect(r.problems).toEqual([]);
    expect([...r.knownGaps]).toEqual(['a/2']);
  });

  test('a check that is not listed needs a fix; a listed check that passes needs a record', () => {
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
        'a/2': ['baseline'],
        'a/3': ['pixels', 'baseline'],
        'a/4': ['pixels'],
      }),
      tools,
      true
    );
    expect(r.problems.map((p) => p.kind)).toEqual(['fix', 'record', 'record']);
    expect(texts(r)[0]).toContain('a/1: fails on pixels 2.00 %');
    expect(texts(r)[1]).toBe('a/3: passes on pixels now');
    expect(texts(r)[2]).toBe('a/4: passes on pixels now');
    expect([...r.knownGaps]).toEqual(['a/2']);
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
      list({ 'b/1': ['pixels'], 'z/1': ['pixels'] }),
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
    const r = checkFixtureGate([deck('a', [slide(1)])], list({ 'z/1': ['pixels'] }), tools, false);
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
    expect(text).toContain("'a/2': ['baseline'],");
    expect(text).toContain('export const AFTER = 1;');
    const mod = (await import(file)) as { FIXTURE_KNOWN_GAPS: KnownGapList };
    expect(mod.FIXTURE_KNOWN_GAPS).toEqual(recorded);
  });

  test('writes the export lane list by its name and leaves the other list as it was', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pptx-gaps-'));
    const file = join(dir, 'thresholds.ts');
    const editor = list({ 'a/1': ['pixels'] });
    writeFileSync(
      file,
      `export interface KnownGapList { slides: object }\n\n${listSource(editor)}\n\n${listSource(list({}), 'FIXTURE_EXPORT_KNOWN_GAPS')}\n`
    );
    const recorded = recordedList([deck('a', [slide(1, baselineFail)])], tools);
    await writeGaps(file, recorded, 'FIXTURE_EXPORT_KNOWN_GAPS');
    const mod = (await import(file)) as {
      FIXTURE_KNOWN_GAPS: KnownGapList;
      FIXTURE_EXPORT_KNOWN_GAPS: KnownGapList;
    };
    expect(mod.FIXTURE_KNOWN_GAPS).toEqual(editor);
    expect(mod.FIXTURE_EXPORT_KNOWN_GAPS).toEqual(recorded);
  });

  test('a partial run, or a deck that did not finish, keeps the entries of the decks it lacks', () => {
    const old = list({ 'a/1': ['pixels'], 'b/1': ['pixels'], 'c/1': ['baseline'] });
    const next = mergeRecorded(
      old,
      [
        deck('a', [slide(1)]),
        deck('b', [], { error: 'the browser crashed while rendering the deck' }),
      ],
      { ...tools, chromium: '2' },
      false
    );
    expect(next.slides).toEqual({ 'b/1': ['pixels'], 'c/1': ['baseline'] });
    // A partial run does not stamp new tools.
    expect(next.chromium).toBe('1');
  });

  test('a complete run drops entries of fixtures that are gone, and stamps new tools', () => {
    const old = list({ 'a/1': ['pixels'], 'gone/1': ['pixels'] }, { a: 4, gone: 3 });
    const next = mergeRecorded(old, [deck('a', [slide(1)])], { ...tools, chromium: '2' }, true);
    expect(next).toMatchObject({ chromium: '2', slides: {}, deckLines: { a: 4 } });
    expect(listChanges(old, next).removed).toEqual(['a/1', 'gone/1', 'gone (its line count)']);
    // A deck error keeps the old tools: the run did not measure every fixture.
    const failed = mergeRecorded(
      old,
      [deck('a', [slide(1)]), deck('gone', [], { error: 'boom' })],
      { ...tools, chromium: '2' },
      true
    );
    expect(failed).toMatchObject({ chromium: '1', slides: { 'gone/1': ['pixels'] } });
  });

  test('listChanges names new kinds, new entries, and removed entries', () => {
    const old = list({ 'a/1': ['pixels'], 'b/1': ['pixels'] });
    const next = list({ 'a/1': ['pixels', 'baseline'], 'c/1': ['pixels'] });
    expect(listChanges(old, next)).toEqual({
      tools: [],
      looser: ['a/1: fails baseline now'],
      added: ["'c/1': ['pixels'],"],
      removed: ['b/1'],
    });
    expect(listChanges(old, list({ 'a/1': ['pixels'] }))).toEqual({
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

  test('--update-gaps removes passing checks, and adds kinds only with --accept-worse', () => {
    const old = list({ 'a/1': ['baseline'] }, { a: 4 });
    const better = [deck('a', [slide(1)])];
    expect(planUpdate(old, better, tools, false, true)).toMatchObject({
      write: true,
      changes: { removed: ['a/1'] },
    });
    const worse = [deck('a', [slide(1, { status: 'fail', baselineDelta: 2, pixelMismatch: 0.1 })])];
    expect(planUpdate(old, worse, tools, false, true).write).toBe(false);
    expect(planUpdate(old, worse, tools, true, true).write).toBe(true);
    const added = [deck('a', [slide(1)]), deck('b', [slide(1, baselineFail)])];
    expect(planUpdate(old, added, tools, false, true)).toMatchObject({
      write: false,
      changes: { added: ["'b/1': ['baseline'],"] },
    });
  });

  test('the committed lists were measured with the tools and fonts in use', () => {
    for (const committed of [FIXTURE_KNOWN_GAPS, FIXTURE_EXPORT_KNOWN_GAPS]) {
      expect(committed.pdfjs).toBe(PDFJS_VERSION);
      expect(committed.fonts).toBe(fixtureFontsIdentity());
      expect(committed.chromium).toBe(pinnedChromium());
    }
  });
});
