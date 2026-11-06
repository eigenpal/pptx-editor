import { describe, expect, test } from 'bun:test';
import {
  changedReferences,
  checkAgainstBase,
  looserGaps,
  looserLimits,
} from './check-known-gaps.ts';
import type { KnownGapList } from './thresholds.ts';
import type { ReferenceManifest } from './reference-identity.ts';

const tools = { pdfjs: '5', chromium: '1', fonts: 'f' };
const list = (slides: KnownGapList['slides'], over: Partial<KnownGapList> = {}): KnownGapList => ({
  ...tools,
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

  test('a smaller or removed entry is allowed', () => {
    expect(looserGaps(base, list({ 'a/1': { pixels: 0.005, lines: 4 } }), new Set())).toEqual([]);
    expect(looserGaps(base, list({}), new Set())).toEqual([]);
  });

  test('a new entry, a new kind, a looser value, or fewer lines is refused', () => {
    expect(
      looserGaps(
        base,
        list({ 'a/1': { pixels: 0.02, baseline: 1, lines: 3 }, 'b/1': { pixels: 0.1 } }),
        new Set()
      )
    ).toEqual([
      'a/1: fewer checked lines (4 → 3)',
      'a/1: a looser pixels score (0.01 → 0.02)',
      'a/1: a new kind, baseline',
      "b/1: a new entry, but the fixture's reference did not change",
    ]);
  });

  test('a changed reference or new tools allow new values', () => {
    const head = list({ 'a/1': { pixels: 0.02, lines: 4 } });
    expect(looserGaps(base, head, new Set(['a']))).toEqual([]);
    expect(looserGaps(base, { ...head, chromium: '2' }, new Set())).toEqual([]);
    // A base from before the PDF comparison has no pdf.js version: every value is new.
    expect(looserGaps({ slides: {} } as unknown as KnownGapList, head, new Set())).toEqual([]);
  });

  test('changedReferences names new fixtures and new decks or PDFs', () => {
    const base: ReferenceManifest = { version: 1, references: { a: identity('1', '2') } };
    const head: ReferenceManifest = {
      version: 1,
      references: { a: identity('1', '3'), b: identity('4', '5') },
    };
    expect([...changedReferences(base, head)]).toEqual(['a', 'b']);
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
