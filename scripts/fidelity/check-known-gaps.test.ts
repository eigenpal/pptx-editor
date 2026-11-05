import { describe, expect, test } from 'bun:test';
import {
  checkAgainstBase,
  looserGaps,
  looserLimits,
  looserThresholds,
  OWNER_GAP_EXCEPTION,
  referenceChanges,
  toolChange,
} from './check-known-gaps.ts';
import { gapEntry } from './known-gaps.ts';
import type { ReferenceIdentity, ReferenceManifest } from './reference-identity.ts';

const manifest = (entries: Record<string, [string, string]>): ReferenceManifest => ({
  version: 1,
  references: Object.fromEntries(
    Object.entries(entries).map(([name, [deck, pdf]]) => [
      name,
      { deck, pdf, pages: 1, pageSizePt: [960, 540] } satisfies ReferenceIdentity,
    ])
  ),
});

const none = { renamed: new Map(), changed: new Map() };

describe('the known-gap ratchet against the base branch', () => {
  const base = { linux: { 'a/1': { pixels: 0.02, baseline: 0.8 } } };

  test('a smaller list or a tighter score passes', () => {
    expect(looserGaps(base, { linux: {} }, none).problems).toEqual([]);
    expect(looserGaps(base, { linux: { 'a/1': { pixels: 0.01 } } }, none).problems).toEqual([]);
  });

  test('a new entry, a new kind, and a looser score fail', () => {
    const head = {
      linux: {
        'a/1': { pixels: 0.03, baseline: 0.8, 'line-breaks': 0.5 },
        'b/1': { pixels: 0.1 },
      },
    };
    expect(looserGaps(base, head, none).problems).toEqual([
      "linux a/1: a looser pixels score (0.02 → 0.03), but the fixture's reference did not change",
      "linux a/1: a new kind, line-breaks, but the fixture's reference did not change",
      "linux b/1: a new entry, but the fixture's reference did not change",
    ]);
  });

  test('a fixture whose reference changed may change its entry, and the change is named', () => {
    const changes = { renamed: new Map(), changed: new Map([['a', 'its deck changed']]) };
    const r = looserGaps(base, { linux: { 'a/1': { pixels: 0.5 } } }, changes);
    expect(r.problems).toEqual([]);
    expect(r.allowed).toEqual([
      'linux a/1: a looser pixels score (0.02 → 0.5), allowed because its deck changed',
    ]);
  });

  test('a renamed fixture keeps the entries of its old name', () => {
    const changes = { renamed: new Map([['b', 'a']]), changed: new Map() };
    expect(looserGaps(base, { linux: { 'b/1': { pixels: 0.02 } } }, changes).problems).toEqual([]);
    expect(looserGaps(base, { linux: { 'b/1': { pixels: 0.9 } } }, changes).problems).toHaveLength(
      1
    );
  });

  test('a new list comes only with its new environment, and every list needs an environment', () => {
    const head = { ...base, windows: { 'a/1': { pixels: 1 } } };
    expect(looserGaps(base, head, none, { base: ['linux'], head: ['linux', 'windows'] })).toEqual({
      problems: [],
      allowed: ['windows: a new list, with its new environment'],
      removed: [],
    });
    expect(
      looserGaps(base, head, none, { base: ['linux', 'windows'], head: ['linux', 'windows'] })
        .problems
    ).toEqual([
      'windows: a new list for an environment that the base already has; add a list only with its new environment',
    ]);
    expect(looserGaps(base, head, none, { base: ['linux'], head: ['linux'] }).problems).toEqual([
      'windows: a list for an environment that FIDELITY_ENVIRONMENTS does not name',
    ]);
    expect(looserGaps(undefined, base, none).problems).toEqual([]);
  });

  test('a list measured again with new tools may take new values, named as ALLOWED', () => {
    const before = {
      linux: { poppler: '24.02.0', chromium: '156', slides: { 'a/1': { pixels: 0.02 } } },
    };
    const after = {
      linux: { poppler: '24.02.0', chromium: '157', slides: { 'a/1': { pixels: 0.03 } } },
    };
    expect(looserGaps(before, after, none)).toEqual({
      problems: [],
      allowed: [
        'linux a/1: a looser pixels score (0.02 → 0.03), allowed because the list was measured again with Chromium 157 (was 156)',
      ],
      removed: [],
    });
  });

  test('a list measured again at a new raster scale may take new values, named as ALLOWED', () => {
    // A base list without `rasterScale` was measured at scale 1.
    const before = {
      linux: { poppler: '24.02.0', chromium: '156', slides: { 'a/1': { pixels: 0.02 } } },
    };
    const after = {
      linux: {
        poppler: '24.02.0',
        chromium: '156',
        rasterScale: 4,
        slides: { 'a/1': { pixels: 0.03 }, 'b/1': { pixels: 0.01 } },
      },
    };
    expect(looserGaps(before, after, none)).toEqual({
      problems: [],
      allowed: [
        'linux a/1: a looser pixels score (0.02 → 0.03), allowed because the list was measured again with raster scale 4 (was 1)',
        'linux b/1: a new entry, allowed because the list was measured again with raster scale 4 (was 1)',
      ],
      removed: [],
    });
    // The same scale on both sides allows nothing.
    const same = { linux: { ...after.linux, slides: { 'a/1': { pixels: 0.04 } } } };
    expect(looserGaps(after, same, none).problems).toEqual([
      "linux a/1: a looser pixels score (0.03 → 0.04), but the fixture's reference did not change",
    ]);
  });

  test('the owner exception allows exactly one looser value', () => {
    const list = (env: string, key: string, pixels: number) => ({
      [env]: { slides: { [key]: { pixels, 'line-breaks': 0, lines: 12 } } },
    });
    const { from, to } = OWNER_GAP_EXCEPTION;
    expect(
      looserGaps(list('linux', 'rtl-text/4', from), list('linux', 'rtl-text/4', to), none)
    ).toEqual({
      problems: [],
      allowed: [
        `linux rtl-text/4: a looser pixels score (${from} → ${to}), allowed once by the owner (OWNER_GAP_EXCEPTION)`,
      ],
      removed: [],
    });
    // Another value, slide, environment, or base value is still a problem.
    const problems = (base: object, head: object) =>
      looserGaps(base as never, head as never, none).problems;
    expect(
      problems(list('linux', 'rtl-text/4', from), list('linux', 'rtl-text/4', 0.0342))
    ).toHaveLength(1);
    expect(
      problems(list('linux', 'rtl-text/3', from), list('linux', 'rtl-text/3', to))
    ).toHaveLength(1);
    expect(
      problems(
        list('macos-reference', 'rtl-text/4', from),
        list('macos-reference', 'rtl-text/4', to)
      )
    ).toHaveLength(1);
    expect(
      problems(list('linux', 'rtl-text/4', 0.0339), list('linux', 'rtl-text/4', to))
    ).toHaveLength(1);
  });

  test('only the one move from raster scale 1 to 4 allows new values', () => {
    const at = (rasterScale: number | undefined, pixels: number) => ({
      linux: {
        poppler: '24.02.0',
        chromium: '156',
        ...(rasterScale === undefined ? {} : { rasterScale }),
        slides: { 'a/1': { pixels } },
      },
    });
    const looser =
      "linux a/1: a looser pixels score (0.03 → 0.04), but the fixture's reference did not change";
    // A later scale change, up or down, re-records nothing.
    expect(looserGaps(at(4, 0.03), at(8, 0.04), none).problems).toEqual([looser]);
    expect(looserGaps(at(4, 0.03), at(1, 0.04), none).problems).toEqual([looser]);
    // Another first move than 1 to 4 allows nothing either.
    expect(looserGaps(at(undefined, 0.03), at(2, 0.04), none).problems).toEqual([looser]);
    expect(toolChange(at(4, 0.03).linux, at(8, 0.04).linux)).toBeUndefined();
  });

  test('only the one move to anti-aliasing judged at the device scale allows new values', () => {
    const at = (antialiasing: string | undefined, pixels: number) => ({
      linux: {
        poppler: '24.02.0',
        chromium: '156',
        rasterScale: 4,
        ...(antialiasing === undefined ? {} : { antialiasing }),
        slides: { 'a/1': { pixels } },
      },
    });
    // A list without the field judged anti-aliasing on the compared images.
    expect(looserGaps(at(undefined, 0.03), at('device', 0.04), none)).toEqual({
      problems: [],
      allowed: [
        'linux a/1: a looser pixels score (0.03 → 0.04), allowed because the list was measured again with anti-aliasing judged on the device images (was images)',
      ],
      removed: [],
    });
    const looser =
      "linux a/1: a looser pixels score (0.03 → 0.04), but the fixture's reference did not change";
    // Back to the compared images, or any other mode, re-records nothing.
    expect(looserGaps(at('device', 0.03), at('images', 0.04), none).problems).toEqual([looser]);
    expect(looserGaps(at('device', 0.03), at('other', 0.04), none).problems).toEqual([looser]);
    expect(looserGaps(at('device', 0.03), at('device', 0.04), none).problems).toEqual([looser]);
  });

  test('a tool change does not allow fewer lines, but a changed deck does', () => {
    const before = {
      linux: { poppler: '24.02.0', chromium: '156', slides: { 'a/1': { pixels: 0.02, lines: 9 } } },
    };
    const after = {
      linux: { poppler: '24.03.0', chromium: '156', slides: { 'a/1': { pixels: 0.03, lines: 7 } } },
    };
    const tools = looserGaps(before, after, none);
    expect(tools.problems).toEqual([
      "linux a/1: fewer checked lines (9 → 7), but the fixture's reference did not change",
    ]);
    expect(tools.allowed).toHaveLength(1);
    const deck = {
      renamed: new Map(),
      changed: new Map([['a', 'its deck changed and has a new reference PDF']]),
      content: new Set(['a']),
    };
    expect(looserGaps(before, after, deck).problems).toEqual([]);
    // A new reference PDF of the same deck allows new scores, but not fewer of the editor's
    // lines.
    const newPdf = {
      renamed: new Map(),
      changed: new Map([['a', 'its reference PDF was replaced']]),
    };
    expect(looserGaps(before, after, newPdf).problems).toEqual([
      "linux a/1: fewer checked lines (9 → 7), but the fixture's reference did not change",
    ]);
  });

  test('fewer checked lines are looser', () => {
    const before = { linux: { 'a/1': { baseline: 1, lines: 9 } } };
    expect(
      looserGaps(before, { linux: { 'a/1': { baseline: 1, lines: 7 } } }, none).problems
    ).toEqual([
      "linux a/1: fewer checked lines (9 → 7), but the fixture's reference did not change",
    ]);
    // An entry that keeps a line kind but loses its line count is looser too.
    expect(looserGaps(before, { linux: { 'a/1': { baseline: 1 } } }, none).problems).toEqual([
      "linux a/1: fewer checked lines (9 → 0), but the fixture's reference did not change",
    ]);
  });

  test('an entry that drops its line kinds keeps its line count, which is still compared', () => {
    const before = {
      linux: { 'a/1': { pixels: 0.02, 'line-breaks': 0.9, baseline: 0.8, lines: 19 } },
    };
    // The slide now passes line breaks and baselines with all its 19 lines: tighter.
    expect(looserGaps(before, { linux: { 'a/1': { pixels: 0.02, lines: 19 } } }, none)).toEqual({
      problems: [],
      allowed: [],
      removed: [],
    });
    // It passes them because 2 lines are gone: not a fix.
    expect(
      looserGaps(before, { linux: { 'a/1': { pixels: 0.02, lines: 17 } } }, none).problems
    ).toEqual([
      "linux a/1: fewer checked lines (19 → 17), but the fixture's reference did not change",
    ]);
  });

  test('a removed entry with a line count is named for the reviewer, never refused', () => {
    // A slide that passes now and a slide whose text is gone both leave the list.
    const before = { linux: { 'a/1': { 'line-breaks': 0.9, lines: 19 }, 'a/2': { pixels: 0.02 } } };
    expect(looserGaps(before, { linux: {} }, none)).toEqual({
      problems: [],
      allowed: [],
      removed: [
        "linux a/1: removed with its 19 checked lines. If the slide still exists, check in this change's fidelity report that it passes with all 19 lines",
      ],
    });
  });

  test('a renamed deck keeps its entries, so no note names them as removed', () => {
    const before = { linux: { 'old/1': { pixels: 0.02, lines: 19 } } };
    const renamed = { renamed: new Map([['new', 'old']]), changed: new Map() };
    expect(
      looserGaps(before, { linux: { 'new/1': { pixels: 0.02, lines: 19 } } }, renamed).removed
    ).toEqual([]);
  });

  test('the scores that replace a size entry are allowed, and named', () => {
    const before = { linux: { 'a/1': { size: true as const } } };
    const after = { linux: { 'a/1': { 'line-breaks': 0.9, lines: 10 } } };
    expect(looserGaps(before, after, none)).toEqual({
      problems: [],
      allowed: ['linux a/1: line-breaks replaces its size entry, which recorded no other score'],
      removed: [],
    });
  });

  test('a size entry measures no lines, so it has no fewer lines', () => {
    const before = { linux: { 'a/1': { pixels: 0.02, lines: 19 } } };
    const gaps = looserGaps(before, { linux: { 'a/1': { size: true } } }, none);
    expect(gaps.problems).toEqual([
      "linux a/1: a new kind, size, but the fixture's reference did not change",
    ]);
  });

  test('the entry that the gate writes for a slide that now passes its line kinds is accepted', () => {
    const before = { linux: { 'a/1': { pixels: 0.02, 'line-breaks': 0.9, lines: 19 } } };
    const now = gapEntry({
      index: 1,
      status: 'fail',
      cause: '',
      diagnostics: [],
      pixelMismatch: 0.02,
      lineMatch: 1,
      linesMatched: 19,
      linesTotal: 19,
      baselineDelta: 0.3,
    });
    expect(now).toEqual({ pixels: 0.02, lines: 19 });
    expect(looserGaps(before, { linux: { 'a/1': now! } }, none).problems).toEqual([]);
  });

  test('a removed list is a problem', () => {
    expect(looserGaps({ ...base, mac: {} }, base, none).problems).toEqual([
      'mac: the list was removed',
    ]);
  });

  test('reads both list shapes: entries alone, and { poppler, chromium, slides }', () => {
    const shaped = {
      linux: { poppler: '24.02.0', chromium: '156', slides: { 'a/1': { pixels: 0.03 } } },
    };
    expect(looserGaps(base, shaped, none).problems).toEqual([
      "linux a/1: a looser pixels score (0.02 → 0.03), but the fixture's reference did not change",
    ]);
  });

  test('a base entry in the kinds-only shape allows any score of its kinds, not a new kind', () => {
    const kindsOnly = { linux: { 'a/1': ['pixels'] as const } };
    expect(looserGaps(kindsOnly, { linux: { 'a/1': { pixels: 0.9 } } }, none).problems).toEqual([]);
    expect(looserGaps(kindsOnly, { linux: { 'a/1': { baseline: 1 } } }, none).problems).toEqual([
      "linux a/1: a new kind, baseline, but the fixture's reference did not change",
    ]);
  });
});

describe('which references really changed', () => {
  test('a new fixture or a changed deck is a change', () => {
    const r = referenceChanges(
      manifest({ a: ['d1', 'p1'], b: ['d2', 'p2'] }),
      manifest({ a: ['d1', 'p1'], b: ['d9', 'p9'], c: ['d4', 'p4'] })
    );
    expect([...r.changed]).toEqual([
      ['b', 'its deck changed and has a new reference PDF'],
      ['c', 'a new fixture'],
    ]);
    expect([...(r.content ?? [])]).toEqual(['b', 'c']);
  });

  test('a changed deck hash with the old PDF is no change: that reference is stale', () => {
    const r = referenceChanges(manifest({ a: ['d1', 'p1'] }), manifest({ a: ['d2', 'p1'] }));
    expect(r.changed.size).toBe(0);
  });

  test('a new reference PDF of the same deck is a change, but not of the line count', () => {
    const r = referenceChanges(manifest({ a: ['d1', 'p1'] }), manifest({ a: ['d1', 'p2'] }));
    expect([...r.changed]).toEqual([['a', 'its reference PDF was replaced']]);
    // The deck is the same, so the editor's own lines are the same.
    expect(r.content?.size).toBe(0);
  });

  test('the same deck and PDF are no change', () => {
    const r = referenceChanges(manifest({ a: ['d1', 'p1'] }), manifest({ a: ['d1', 'p1'] }));
    expect(r.changed.size).toBe(0);
  });

  test('the same deck bytes under a new name is a rename, not a new fixture', () => {
    const r = referenceChanges(manifest({ a: ['d1', 'p1'] }), manifest({ b: ['d1', 'p1'] }));
    expect([...r.renamed]).toEqual([['b', 'a']]);
    expect(r.changed.size).toBe(0);
  });
});

describe('the thresholds against the base', () => {
  test('the raster width, the page tolerance, and the margins only tighten', () => {
    expect(looserThresholds({ renderWidthPx: 1920 }, { renderWidthPx: 960 })).toEqual([
      'THRESHOLDS.renderWidthPx is looser (1920 → 960)',
    ]);
    expect(
      looserLimits(
        { page: 2, margins: { pixels: 0.0001 } },
        { page: 3, margins: { pixels: 0.001 } }
      )
    ).toEqual([
      'PAGE_SIZE_TOLERANCE_PT is looser (2 → 3)',
      'GAP_MARGINS.pixels is looser (0.0001 → 0.001)',
    ]);
    expect(looserLimits({}, { page: 2, margins: { pixels: 0.0001 } })).toEqual([]);
  });

  test('the raster scale changes only once, from 1 to 4; a base without it compared at scale 1', () => {
    const only = 'only the one move from 1 to 4 is allowed';
    expect(looserThresholds({}, { rasterScale: 4 })).toEqual([]);
    expect(looserThresholds({ rasterScale: 1 }, { rasterScale: 4 })).toEqual([]);
    expect(looserThresholds({ rasterScale: 4 }, { rasterScale: 4 })).toEqual([]);
    expect(looserThresholds({ rasterScale: 4 }, { rasterScale: 2 })).toEqual([
      `THRESHOLDS.rasterScale changed (4 → 2); ${only}`,
    ]);
    expect(looserThresholds({ rasterScale: 4 }, { rasterScale: 8 })).toEqual([
      `THRESHOLDS.rasterScale changed (4 → 8); ${only}`,
    ]);
    expect(looserThresholds({}, { rasterScale: 2 })).toEqual([
      `THRESHOLDS.rasterScale changed (1 → 2); ${only}`,
    ]);
  });

  test('the anti-aliasing mode changes only once, from images to device', () => {
    const only = 'only the one move from images to device is allowed';
    expect(looserThresholds({}, { antialiasing: 'device' })).toEqual([]);
    expect(looserThresholds({ antialiasing: 'device' }, { antialiasing: 'device' })).toEqual([]);
    expect(looserThresholds({ antialiasing: 'device' }, { antialiasing: 'images' })).toEqual([
      `THRESHOLDS.antialiasing changed (device → images); ${only}`,
    ]);
    expect(looserThresholds({ antialiasing: 'device' }, {})).toEqual([
      `THRESHOLDS.antialiasing changed (device → images); ${only}`,
    ]);
  });

  test('only tighten', () => {
    const base = { maxPixelMismatch: 0.005, minLineBreakMatch: 1, maxBaselineDeltaPt: 0.5 };
    expect(looserThresholds(base, { ...base, maxPixelMismatch: 0.004 })).toEqual([]);
    expect(
      looserThresholds(base, { ...base, maxPixelMismatch: 0.01, minLineBreakMatch: 0.9 })
    ).toEqual([
      'THRESHOLDS.maxPixelMismatch is looser (0.005 → 0.01)',
      'THRESHOLDS.minLineBreakMatch is looser (1 → 0.9)',
    ]);
    expect(looserThresholds(undefined, base)).toEqual([]);
  });
});

describe('checkAgainstBase', () => {
  test('this checkout against its own HEAD loosens nothing', async () => {
    const r = await checkAgainstBase('HEAD');
    expect(r.exitCode).toBe(0);
    expect(r.lines.at(-1)).toBe(
      'Nothing is added or loosened against HEAD, except what ALLOWED lines name.'
    );
  });

  test('a ref that does not exist is a usage error, never a pass', async () => {
    const r = await checkAgainstBase('origin/no-such-branch-for-a-test');
    expect(r).toEqual({
      lines: ['origin/no-such-branch-for-a-test is not a commit. Fetch the base branch first.'],
      exitCode: 2,
    });
  });
});
