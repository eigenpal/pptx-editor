import { describe, expect, test } from 'bun:test';
import {
  checkAgainstBase,
  looserGaps,
  looserLimits,
  newFixtures,
  nonSourceFiles,
  toolOnlyLockChange,
  toolOnlyManifestChange,
} from './check-known-gaps.ts';
import type { KnownGapList } from './thresholds.ts';
import type { ReferenceManifest } from './reference-identity.ts';

const tools = { pdfjs: '5', chromium: '1', fonts: 'f' };
const list = (slides: KnownGapList['slides'], over: Partial<KnownGapList> = {}): KnownGapList => ({
  ...tools,
  deckLines: {},
  slides,
  ...over,
});
const identity = (deck: string, pdf: string) => ({
  deck,
  pdf,
  pages: 1,
  pageSizePt: [960, 540] as const,
});

describe('the known gaps only shrink', () => {
  const base = list({ 'a/1': { pixels: 0.01, lines: 4 } });

  const source = ['packages/core/src/output/paint.ts'];

  test('a smaller or removed entry is allowed', () => {
    expect(
      looserGaps(base, list({ 'a/1': { pixels: 0.005, lines: 4 } }), new Set(), source)
    ).toEqual([]);
    expect(looserGaps(base, list({}), new Set(), source)).toEqual([]);
  });

  test('with a source change, a new entry, a new kind, a looser value, or fewer lines is refused', () => {
    expect(
      looserGaps(
        base,
        list({ 'a/1': { pixels: 0.02, baseline: 1, lines: 3 }, 'b/1': { pixels: 0.1 } }),
        new Set(),
        source
      )
    ).toEqual([
      'a/1: fewer checked lines (4 → 3)',
      'a/1: a looser pixels score (0.01 → 0.02)',
      'a/1: a new kind, baseline',
      'b/1: a new entry',
      'the change also touches source (packages/core/src/output/paint.ts): record looser entries in a pull request that changes only tools, fixtures, or references',
    ]);
  });

  test('a change with no source, or a new fixture, may loosen', () => {
    const head = list({ 'a/1': { pixels: 0.02, lines: 4 } });
    expect(looserGaps(base, head, new Set(), [])).toEqual([]);
    expect(looserGaps(base, head, new Set(['a']), source)).toEqual([]);
    // A base from before the PDF comparison has no pdf.js version: every value is new.
    expect(looserGaps({ slides: {} } as unknown as KnownGapList, head, new Set(), source)).toEqual(
      []
    );
  });

  test('fewer checked lines in a deck are refused', () => {
    const head = list({ 'a/1': { pixels: 0.01, lines: 4 } }, { deckLines: { a: 10 } });
    expect(looserGaps({ ...base, deckLines: { a: 12 } }, head, new Set(), source)).toEqual([
      'a: checks fewer lines (12 → 10)',
      'the change also touches source (packages/core/src/output/paint.ts): record looser entries in a pull request that changes only tools, fixtures, or references',
    ]);
  });

  test('newFixtures names only the fixtures that the base lacks', () => {
    const base: ReferenceManifest = { version: 1, references: { a: identity('1', '2') } };
    const head: ReferenceManifest = {
      version: 1,
      references: { a: identity('9', '3'), b: identity('4', '5') },
    };
    expect([...newFixtures(base, head)]).toEqual(['b']);
  });

  test('only versions of the tool packages count as a tools change', () => {
    const pkg = (deps: Record<string, string>, extra = {}) =>
      JSON.stringify({ name: 'x', devDependencies: deps, ...extra });
    const base = pkg({ 'pdfjs-dist': '5.4.624', react: '19.2.0' });
    expect(toolOnlyManifestChange(base, pkg({ 'pdfjs-dist': '5.5.0', react: '19.2.0' }))).toBe(
      true
    );
    expect(toolOnlyManifestChange(base, pkg({ 'pdfjs-dist': '5.4.624', react: '19.3.0' }))).toBe(
      false
    );
    expect(
      toolOnlyManifestChange(
        base,
        pkg({ 'pdfjs-dist': '5.4.624', react: '19.2.0' }, { scripts: {} })
      )
    ).toBe(false);
    expect(toolOnlyManifestChange(undefined, base)).toBe(false);
    const lock = (pdfjs: string, canvas: string, react: string) =>
      `{
  "lockfileVersion": 1,
  "workspaces": { "": { "devDependencies": { "pdfjs-dist": "${pdfjs}", "react": "^19.2.0", }, }, },
  "packages": {
    "pdfjs-dist": ["pdfjs-dist@${pdfjs}", "", { "optionalDependencies": { "@napi-rs/canvas": "^${canvas}" } }, "sha"],
    "@napi-rs/canvas": ["@napi-rs/canvas@${canvas}", "", {}, "sha"],
    "react": ["react@${react}", "", {}, "sha"],
  }
}`;
    const old = lock('5.4.624', '0.1.80', '19.2.0');
    // pdf.js and the package it depends on.
    expect(toolOnlyLockChange(old, lock('5.5.0', '0.1.81', '19.2.0'))).toBe(true);
    // An engine dependency.
    expect(toolOnlyLockChange(old, lock('5.4.624', '0.1.80', '19.2.1'))).toBe(false);
    const read = (file: string) =>
      file.endsWith('package.json')
        ? { base, head: pkg({ 'pdfjs-dist': '5.5.0', react: '19.2.0' }) }
        : { base: old, head: lock('5.4.624', '0.1.80', '19.2.1') };
    expect(
      nonSourceFiles(
        [
          'bun.lock',
          'package.json',
          'packages/core/package.json',
          'scripts/fidelity/thresholds.ts',
          'fixtures/fonts/New.ttf',
          'fixtures/tables.pdf',
          'fixtures/references.json',
          'scripts/fixtures/build.ts',
          'scripts/fidelity/compare.ts',
          'packages/core/src/output/paint.ts',
        ],
        read
      )
    ).toEqual(['bun.lock', 'scripts/fidelity/compare.ts', 'packages/core/src/output/paint.ts']);
  });

  test('limits only tighten', () => {
    const base = { renderWidthPx: 1920, maxPixelMismatch: 0.005, minLineBreakMatch: 1 };
    expect(looserLimits(base, { ...base, maxPixelMismatch: 0.004 })).toEqual([]);
    expect(
      looserLimits(base, { renderWidthPx: 960, maxPixelMismatch: 0.01, minLineBreakMatch: 0.9 })
    ).toEqual([
      'renderWidthPx is looser (1920 → 960)',
      'maxPixelMismatch is looser (0.005 → 0.01)',
      'minLineBreakMatch is looser (1 → 0.9)',
    ]);
  });

  test('a limit that goes away is reported', () => {
    const base = { maxPixelMismatch: 0.005, maxSizeDeltaPx: 0 };
    expect(looserLimits(base, { maxPixelMismatch: 0.005 }, {})).toEqual([
      'maxSizeDeltaPx was removed (it was 0); list it in RETIRED_LIMITS with the reason',
    ]);
    expect(looserLimits(base, { maxPixelMismatch: 0.005 }, { maxSizeDeltaPx: 'a reason' })).toEqual(
      []
    );
  });
});

describe('checkAgainstBase', () => {
  test('this checkout against its own HEAD loosens nothing', async () => {
    const r = await checkAgainstBase('HEAD');
    expect(r.exitCode).toBe(0);
  });

  test('a ref that does not exist is a usage error, never a pass', async () => {
    const r = await checkAgainstBase('origin/no-such-branch-for-a-test');
    expect(r).toEqual({
      lines: ['origin/no-such-branch-for-a-test is not a commit. Fetch the base branch first.'],
      exitCode: 2,
    });
  });
});
