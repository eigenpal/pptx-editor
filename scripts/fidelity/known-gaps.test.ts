import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  checkFixtureGate,
  failureKinds,
  fidelityEnvironment,
  gapEntry,
  gapEntryLine,
  gapKinds,
  gateExitCode,
  listToolChanges,
  recordedGaps,
  slideVerdict,
} from './known-gaps.ts';
import { deckFacts } from './reference-identity.ts';
import { sizeLimitText } from './compare.ts';
import {
  FIDELITY_ENVIRONMENTS,
  FIXTURE_KNOWN_GAPS,
  THRESHOLDS,
  type KnownGap,
} from './thresholds.ts';
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

const baselineFail = { status: 'fail' as const, baselineDelta: 0.9 };
const gaps: Record<string, KnownGap> = { 'a/2': { baseline: 0.9, lines: 4 } };
const complete = { complete: true, list: 'FIXTURE_KNOWN_GAPS.test' };

describe('failureKinds', () => {
  test('names each threshold a failing slide misses', () => {
    expect(failureKinds(slide(1))).toEqual([]);
    expect(
      failureKinds(
        slide(1, { status: 'fail', pixelMismatch: 0.02, lineMatch: 0.5, baselineDelta: 3 })
      )
    ).toEqual(['pixels', 'line-breaks', 'baseline']);
    expect(
      failureKinds(slide(1, { status: 'fail', sizeMismatch: true, pixelMismatch: 1 }))
    ).toEqual(['size']);
    expect(failureKinds(slide(1, { status: 'stale' }))).toEqual([]);
  });

  test('a slide with no checked line misses no line threshold', () => {
    const s = slide(1, { status: 'fail', pixelMismatch: 0.02, lineMatch: undefined });
    expect(failureKinds({ ...s, baselineDelta: undefined })).toEqual(['pixels']);
  });
});

describe('gapEntry', () => {
  test('records the measured value of each missed threshold', () => {
    const s = slide(1, {
      status: 'fail',
      pixelMismatch: 0.08963,
      lineMatch: 2 / 3,
      baselineDelta: 0.6831,
    });
    expect(gapEntry(s)).toEqual({
      pixels: 0.0896,
      'line-breaks': 0.6667,
      baseline: 0.68,
      lines: 4,
    });
    expect(gapEntry(slide(1))).toBeUndefined();
    expect(gapEntry(slide(1, { status: 'fail', sizeMismatch: true }))).toEqual({ size: true });
  });

  test('records the checked lines also when only the pixels fail', () => {
    // So an entry that drops its line kinds keeps its count, and lost text cannot pass as a fix.
    const s = slide(1, { status: 'fail', pixelMismatch: 0.02, lineMatch: 1, linesTotal: 7 });
    expect(gapEntry(s)).toEqual({ pixels: 0.02, lines: 7 });
    const noText = { linesTotal: 0, linesMatched: 0, lineMatch: undefined };
    expect(gapEntry(slide(1, { status: 'fail', pixelMismatch: 0.02, ...noText }))).toEqual({
      pixels: 0.02,
    });
  });

  test('prints an entry as a line of thresholds.ts', () => {
    expect(gapEntryLine('a/1', { pixels: 0.01, 'line-breaks': 0.5 })).toBe(
      "'a/1': { pixels: 0.01, 'line-breaks': 0.5 },"
    );
    expect(gapKinds({ baseline: 1, size: true })).toEqual(['size', 'baseline']);
  });
});

describe('checkFixtureGate', () => {
  test('passes when every failure is listed at its recorded value, and marks it a known gap', () => {
    const r = checkFixtureGate([deck('a', [slide(1), slide(2, baselineFail)])], gaps, complete);
    expect(r.problems).toEqual([]);
    expect([...r.knownGaps]).toEqual(['a/2']);
  });

  test('a change within the noise margin is the same value', () => {
    const near = slide(2, { ...baselineFail, baselineDelta: 0.905 });
    expect(checkFixtureGate([deck('a', [slide(1), near])], gaps, complete).problems).toEqual([]);
  });

  test('fails on a new failing slide and on a new kind for a listed slide', () => {
    const r = checkFixtureGate(
      [
        deck('a', [
          slide(1, { status: 'fail', pixelMismatch: 0.02 }),
          slide(2, { ...baselineFail, lineMatch: 0.5, linesMatched: 2 }),
        ]),
      ],
      gaps,
      complete
    );
    expect(r.problems).toEqual([
      "a/1: fails on pixels 2.00 % (threshold ≤ 0.50 %), which is not a known gap. Fix the slide; only a new fixture, a changed deck, or a changed environment may add 'a/1': { pixels: 0.02, lines: 4 },",
      "a/2: fails on line breaks 2 of 4 (50.00 %) (threshold ≥ 100.00 %), which is not a known gap. Fix the slide; only a new fixture, a changed deck, or a changed environment may add 'a/2': { 'line-breaks': 0.5, baseline: 0.9, lines: 4 },",
    ]);
    expect(r.knownGaps.size).toBe(0);
  });

  test('fails when a listed slide scores worse than its recorded value', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1), slide(2, { ...baselineFail, baselineDelta: 15 })])],
      gaps,
      complete
    );
    expect(r.problems).toEqual([
      'a/2: baseline 15.00 pt is worse than its known gap (baseline 0.90 pt). Fix the regression: a known gap does not loosen.',
    ]);
    expect(r.knownGaps.size).toBe(0);
  });

  test('fails when a listed slide scores better, until the change records the better value', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1), slide(2, { ...baselineFail, baselineDelta: 0.6 })])],
      gaps,
      complete
    );
    expect(r.problems).toEqual([
      "a/2: scores better than its known gap on baseline. Record it in this change: in FIXTURE_KNOWN_GAPS.test, set 'a/2': { baseline: 0.6, lines: 4 },",
    ]);
  });

  test('fails when a listed failure is gone, and prints the entry to use', () => {
    expect(checkFixtureGate([deck('a', [slide(1), slide(2)])], gaps, complete).problems).toEqual([
      "a/2: passes on baseline now. Record it in this change: in FIXTURE_KNOWN_GAPS.test, remove 'a/2'",
    ]);
    const two: Record<string, KnownGap> = { 'a/1': { pixels: 0.02, baseline: 2 } };
    const r = checkFixtureGate(
      [deck('a', [slide(1, { status: 'fail', pixelMismatch: 0.02 })])],
      two,
      complete
    );
    expect(r.problems).toEqual([
      "a/1: passes on baseline now. Record it in this change: in FIXTURE_KNOWN_GAPS.test, set 'a/1': { pixels: 0.02, lines: 4 },",
    ]);
  });

  test('a listed line kind that is no longer measured is a problem, never "passes now"', () => {
    const listed: Record<string, KnownGap> = { 'a/1': { 'line-breaks': 0.5, baseline: 1 } };
    // The slide's text disappeared: no checked line, so no line score at all.
    const gone = slide(1, {
      lineMatch: undefined,
      linesTotal: 0,
      linesMatched: 0,
      baselineDelta: undefined,
    });
    const r = checkFixtureGate([deck('a', [gone])], listed, complete);
    expect(r.problems).toEqual([
      'a/1: its line breaks are no longer measured: the slide has no checked line now (0 of 0 lines). Find where the text went; a known gap does not end this way',
      'a/1: its baseline is no longer measured: the slide has no matched line now (0 of 0 lines). Find where the text went; a known gap does not end this way',
    ]);
  });

  test('a listed slide that checks fewer lines than recorded fails, even if its scores look better', () => {
    const listed: Record<string, KnownGap> = { 'a/1': { 'line-breaks': 0.6667, lines: 3 } };
    // One unmatched line stopped rendering: 2 of 2 lines now, which looks like a pass.
    const partial = slide(1, {
      lineMatch: 1,
      linesMatched: 2,
      linesTotal: 2,
      pdfLinesUnmatched: 1,
    });
    expect(checkFixtureGate([deck('a', [partial])], listed, complete).problems).toEqual([
      "a/1: checks 2 of its 3 recorded lines now. Find where the text went; a known gap does not end this way. Only if this change edits the deck (or exports it with another font set), record it: in FIXTURE_KNOWN_GAPS.test, remove 'a/1'",
    ]);
  });

  test('names a tool change in a changed score, and notes it', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1), slide(2, { ...baselineFail, baselineDelta: 2 })])],
      gaps,
      { ...complete, toolChanges: ['Poppler 26.11.0 (the list: 26.10.0)'] }
    );
    expect(r.notes).toEqual([
      'this run uses other tools than FIXTURE_KNOWN_GAPS.test was measured with: Poppler 26.11.0 (the list: 26.10.0)',
    ]);
    expect(r.problems[0]).toContain('If only the tools changed, record the list again for them.');
  });

  test('another line count fails until it is recorded, also on a pixels entry', () => {
    const pixels = slide(1, { status: 'fail', pixelMismatch: 0.02 });
    for (const [entry, recorded] of [
      [{ pixels: 0.02, lines: 3 }, '3'],
      [{ pixels: 0.02 }, 'no line count'],
    ] as const) {
      const r = checkFixtureGate([deck('a', [pixels])], { 'a/1': entry }, complete);
      expect(r.problems).toEqual([
        `a/1: checks 4 lines, and its entry records ${recorded}. Record it in this change: in FIXTURE_KNOWN_GAPS.test, set 'a/1': { pixels: 0.02, lines: 4 },`,
      ]);
      expect(r.knownGaps.size).toBe(0);
    }
    const same = checkFixtureGate(
      [deck('a', [pixels])],
      { 'a/1': { pixels: 0.02, lines: 4 } },
      complete
    );
    expect(same.problems).toEqual([]);
  });

  test('a fixed size asks to record the scores that replace it, never as a new failure', () => {
    const fixed = slide(1, { status: 'fail', pixelMismatch: 0.02 });
    const r = checkFixtureGate([deck('a', [fixed])], { 'a/1': { size: true } }, complete);
    expect(r.problems).toEqual([
      "a/1: passes on size now. Record it in this change: in FIXTURE_KNOWN_GAPS.test, set 'a/1': { pixels: 0.02, lines: 4 },",
    ]);
    // A fixed size on a slide that passes everything leaves the list.
    const clean = checkFixtureGate([deck('a', [slide(1)])], { 'a/1': { size: true } }, complete);
    expect(clean.problems).toEqual([
      "a/1: passes on size now. Record it in this change: in FIXTURE_KNOWN_GAPS.test, remove 'a/1'",
    ]);
  });

  test('a listed size failure is a known gap, with no line count to record', () => {
    const sized = slide(1, { status: 'fail', sizeMismatch: true, pixelMismatch: undefined });
    const r = checkFixtureGate([deck('a', [sized])], { 'a/1': { size: true } }, complete);
    expect(r.problems).toEqual([]);
    expect([...r.knownGaps]).toEqual(['a/1']);
  });

  test('a size failure with a line cause keeps the sizes on the size cause', () => {
    const cause =
      'slide size differs: the render does not have its expected size; line breaks differ';
    const sizeDetail = 'the render is 1920 × 1300 px, 1920 × 1358 px expected';
    const failed = slide(1, {
      status: 'fail',
      sizeMismatch: true,
      pixelMismatch: undefined,
      cause,
      sizeDetail,
    });
    const r = checkFixtureGate([deck('a', [failed])], {}, complete);
    expect(r.problems[0]).toStartWith(
      `a/1: slide size differs: the render does not have its expected size — ${sizeDetail} (size check`
    );
  });

  test('a size mismatch hides the other scores, so they are not "gone"', () => {
    const listed: Record<string, KnownGap> = { 'a/1': { pixels: 0.02 } };
    const cause = 'slide size differs: the render does not have its expected size';
    const sizeDetail =
      'the render is 1920 × 1300 px, 1920 × 1358 px expected, and the PDF raster is 1920 × 1358 px, as expected';
    const r = checkFixtureGate(
      [
        deck('a', [
          slide(1, {
            status: 'fail',
            sizeMismatch: true,
            pixelMismatch: undefined,
            cause,
            sizeDetail,
          }),
        ]),
      ],
      listed,
      complete
    );
    expect(r.problems).toEqual([
      `a/1: ${cause} — ${sizeDetail} (size check ${sizeLimitText()}), which is not a known gap. Find which size is wrong: the render size comes from renderSizePx and examples/vite/src/harness.ts, the raster size from expectedRasterSizePx (CONTRIBUTING.md, "Fidelity harness"). Only a new fixture, a changed deck, or a changed environment may add 'a/1': { size: true },`,
    ]);
  });

  test('fails on a missing or stale reference with its reasons, an error, and a save difference', () => {
    const r = checkFixtureGate(
      [
        deck('a', [slide(1, { status: 'stale' }), slide(2, { status: 'stale' })], {
          reference: 'stale',
          referenceIssues: ['the deck changed after its reference was recorded'],
        }),
        deck('b', [slide(1, { status: 'no-reference', selfConsistent: false })], {
          reference: 'none',
          referenceIssues: ['no reference PDF'],
        }),
        deck('c', [], { error: 'the deck does not open: zip' }),
      ],
      gaps,
      { complete: false }
    );
    expect(r.problems).toEqual([
      'a: no usable reference (the deck changed after its reference was recorded). It needs a reference PDF next to it and its entry in references.json.',
      'b: no usable reference (no reference PDF). It needs a reference PDF next to it and its entry in references.json.',
      'b/1: save → reopen → render differs from the first render',
      'c: the harness could not score the deck (the deck does not open: zip)',
    ]);
  });

  test('fails on a slide that was not compared although the deck has a current PDF', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1, { status: 'no-reference', cause: 'the PDF raster has no page 1' })])],
      undefined,
      { complete: true }
    );
    expect(r.problems).toEqual(['a/1: not compared (the PDF raster has no page 1)']);
  });

  test('a hidden slide has no PDF page and is not a problem', () => {
    const hidden = slide(1, { status: 'no-reference', hidden: true });
    expect(checkFixtureGate([deck('a', [hidden])], undefined, complete).problems).toEqual([]);
  });

  test('fails on a failing slide that misses no known threshold', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1, { status: 'fail', pixelMismatch: Number.NaN })])],
      undefined,
      complete
    );
    expect(r.problems).toEqual(['a/1: fails, but on no threshold that the gate knows']);
  });

  test('a listed slide that does not exist fails a complete run only', () => {
    const run = [deck('a', [slide(1)])];
    expect(checkFixtureGate(run, gaps, { complete: false }).problems).toEqual([]);
    expect(checkFixtureGate(run, gaps, complete).problems).toEqual([
      'a/2: listed in FIXTURE_KNOWN_GAPS.test, but no such fixture slide exists',
    ]);
  });

  test('a deck that does not open reports once, not once per listed slide', () => {
    const r = checkFixtureGate([deck('a', [], { error: 'zip' })], gaps, complete);
    expect(r.problems).toEqual(['a: the harness could not score the deck (zip)']);
  });

  test('without a list for the environment, failures are reported, not gated', () => {
    const r = checkFixtureGate(
      [deck('a', [slide(1, { status: 'fail', pixelMismatch: 0.5 })])],
      undefined,
      complete
    );
    expect(r.problems).toEqual([]);
  });
});

describe('the report lines', () => {
  test('a failing slide names its scores, the thresholds, and its known gap', () => {
    const s = slide(2, baselineFail);
    expect(slideVerdict(s, gaps['a/2'])).toBe(
      'fails: baseline 0.90 pt (threshold ≤ 0.5 pt). Known gap: baseline 0.90 pt'
    );
    expect(slideVerdict(slide(1), undefined)).toBe('');
  });

  test('recordedGaps lists the entries of every failing slide', () => {
    expect(recordedGaps([deck('a', [slide(1), slide(2, baselineFail)])])).toEqual([
      "'a/2': { baseline: 0.9, lines: 4 },",
    ]);
  });
});

describe('fidelityEnvironment', () => {
  test('CI and macOS with the reference fonts have lists; other machines do not', () => {
    expect(fidelityEnvironment('linux', false)).toBe('linux');
    expect(fidelityEnvironment('darwin', true)).toBe('macos-reference');
    expect(fidelityEnvironment('darwin', false)).toBeUndefined();
    expect(fidelityEnvironment('win32', false)).toBeUndefined();
  });
});

describe('FIXTURE_KNOWN_GAPS', () => {
  test('there is one list for each environment, with its tool versions', () => {
    expect(Object.keys(FIXTURE_KNOWN_GAPS).sort()).toEqual([...FIDELITY_ENVIRONMENTS].sort());
    for (const list of Object.values(FIXTURE_KNOWN_GAPS)) {
      expect(list.poppler).toMatch(/^\d+\.\d+/);
      expect(list.chromium).toMatch(/^\d+\.\d+/);
    }
  });

  test('every list was measured at the raster scale that the harness uses', () => {
    for (const list of Object.values(FIXTURE_KNOWN_GAPS))
      expect(list.rasterScale).toBe(THRESHOLDS.rasterScale);
    for (const list of Object.values(FIXTURE_KNOWN_GAPS))
      expect(list.antialiasing).toBe(THRESHOLDS.antialiasing);
  });

  test('every entry names an existing fixture slide and at least one kind', () => {
    const fixtures = join(import.meta.dir, '../../fixtures');
    const wrong: string[] = [];
    for (const [env, list] of Object.entries(FIXTURE_KNOWN_GAPS))
      for (const [key, gap] of Object.entries(list.slides)) {
        const m = /^([a-z0-9-]+)\/(\d+)$/.exec(key);
        const file = m ? join(fixtures, `${m[1]}.pptx`) : '';
        const slides =
          m && existsSync(file) ? deckFacts(new Uint8Array(readFileSync(file))).slides : 0;
        if (!m || Number(m[2]) < 1 || Number(m[2]) > slides || gapKinds(gap).length === 0)
          wrong.push(`${env}: ${key}`);
      }
    expect(wrong).toEqual([]);
  });

  test('every value is a measure: pixels and line breaks are fractions, baselines positive', () => {
    const wrong: string[] = [];
    for (const [env, list] of Object.entries(FIXTURE_KNOWN_GAPS))
      for (const [key, gap] of Object.entries(list.slides)) {
        const fraction = (v: number | undefined): boolean => v === undefined || (v >= 0 && v <= 1);
        if (!fraction(gap.pixels) || !fraction(gap['line-breaks'])) wrong.push(`${env}: ${key}`);
        if (gap.baseline !== undefined && !(gap.baseline > 0)) wrong.push(`${env}: ${key}`);
        // A line kind needs its line count; a pixels entry may have one too (`gapEntry`).
        const lineKind = gap['line-breaks'] !== undefined || gap.baseline !== undefined;
        if ((lineKind && gap.lines === undefined) || (gap.lines !== undefined && !(gap.lines > 0)))
          wrong.push(`${env}: ${key} lines`);
      }
    expect(wrong).toEqual([]);
  });
});

describe('the gate result', () => {
  test('names each tool that differs from the list', () => {
    const recorded = { poppler: '24.02.0', chromium: '156' };
    expect(listToolChanges(recorded, '24.02.0', '156')).toEqual([]);
    expect(listToolChanges(recorded, '26.10.0', '157')).toEqual([
      'Poppler 26.10.0 (the list: 24.02.0)',
      'Chromium 157 (the list: 156)',
    ]);
    expect(listToolChanges(undefined, '1', '2')).toEqual([]);
  });

  test('exits 1 with problems, 3 when nothing was gated, else 0', () => {
    expect(gateExitCode(2, true)).toBe(1);
    expect(gateExitCode(2, false)).toBe(1);
    expect(gateExitCode(0, false)).toBe(3);
    expect(gateExitCode(0, true)).toBe(0);
  });
});
