import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  approximationTitles,
  otherDiagnostics,
  probableCause,
  undetectedApproximations,
  writeReport,
  type SlideResult,
} from './report.ts';

const failing: SlideResult = {
  index: 1,
  render: 'slide1.png',
  status: 'fail',
  pixelMismatch: 0.02,
  lineMatch: 1,
  baselineDelta: 0.1,
  diagnostics: [],
  cause: '',
};

describe('fidelity report', () => {
  test('an unrendered effect is a probable cause', () => {
    expect(probableCause(failing, ['effect not rendered: presetShdw'])).toBe(
      'effect not rendered (presetShdw)'
    );
  });

  test('the approximations a slide uses are probable causes, by their docs title', () => {
    const diagnostics = ['approximation: blur', 'approximation: glow-soft-edge'];
    expect(probableCause(failing, diagnostics)).toBe('approximations: Blur, Glow and soft edge');
    expect(approximationTitles(diagnostics)).toEqual(['Blur', 'Glow and soft edge']);
  });

  test('a line that differs only by its bullet or number is a limit of the line check', () => {
    const bullets: SlideResult = {
      ...failing,
      lineMatch: 0.5,
      unmatchedTexts: ['First level', 'Child a'],
      pdfUnmatchedTexts: ['• First level', 'a. Child a'],
    };
    expect(probableCause(bullets, [])).toBe(
      'bullet or number: only the PDF line holds it (a limit of the line check)'
    );
    const wrapped: SlideResult = { ...bullets, pdfUnmatchedTexts: ['• First level', 'Child'] };
    expect(probableCause(wrapped, [])).toBe('line breaks or text differ from the PDF');
  });

  test('a line that the PDF splits into other words, or does not hold, says so', () => {
    const lines = (ours: string[], pdf: string[]): SlideResult => ({
      ...failing,
      pixelMismatch: 0,
      lineMatch: 0.5,
      unmatchedTexts: ours,
      pdfUnmatchedTexts: pdf,
    });
    // `compareLines` found each unmatched line split into words on its row (`text-runs`, `spacing`).
    expect(
      probableCause(
        {
          ...lines(['small LARGE line height'], ['small', 'LARGE line height']),
          lineSplitOnly: true,
        },
        []
      )
    ).toBe('the PDF splits the same text into other words (a limit of the line check)');
    // `autofit`: the same words, broken at other places, is a real line-break difference.
    expect(
      probableCause(
        lines(
          ['The quick brown fox jumps', 'over the lazy dog'],
          ['The quick brown fox', 'jumps over the lazy dog']
        ),
        []
      )
    ).toBe('line breaks differ');
    // `fills-lines` slide 1: the PDF has no text layer for the line.
    expect(probableCause(lines(['Gradient text'], []), [])).toBe(
      'the PDF has no text for the laid-out line'
    );
    // `preserve-unsupported`: the reference PDF draws another branch.
    expect(probableCause(lines(['Fallback branch'], ['Choice branch']), [])).toBe(
      'line breaks or text differ from the PDF'
    );
  });

  test('a slide with no known cause points to the diff area, and guesses no cause', () => {
    expect(probableCause(failing, [])).toBe(
      'pixels differ, and no layout diagnostic names a cause: see the diff area'
    );
  });

  test('the line model is a cause only when a line check fails', () => {
    const d = ['approximation: line-model'];
    expect(probableCause(failing, d)).toBe(
      'pixels differ, and no layout diagnostic names a cause: see the diff area'
    );
    expect(probableCause({ ...failing, baselineDelta: 0.6, estimatedBaselines: 0 }, d)).toBe(
      'line height or baseline model; approximations: Line model'
    );
  });

  test('approximations and other diagnostics are listed apart', () => {
    const d = ['font substituted: Aptos', 'approximation: line-model'];
    expect(otherDiagnostics(d)).toEqual(['font substituted: Aptos']);
    expect(approximationTitles(d)).toEqual(['Line model']);
  });

  test('approximations layout cannot detect are listed for every report', () => {
    expect(undetectedApproximations()).toContain('Color transform spaces');
    expect(undetectedApproximations()).not.toContain('Blur');
  });
});

test('the report shows a size failure as sizes, with no pixel score and no probable cause', () => {
  const out = mkdtempSync(join(tmpdir(), 'pptx-report-'));
  try {
    const sized: SlideResult = {
      index: 1,
      status: 'fail',
      diagnostics: [],
      cause: 'slide size differs: the render does not have its expected size',
      sizeMismatch: true,
      renderSizePx: [1920, 1300],
      referenceSizePx: [1920, 1358],
      expectedRenderSizePx: [1920, 1358],
      expectedReferenceSizePx: [1920, 1358],
      sizeDetail: 'the render is 1920 × 1300 px, 1920 × 1358 px expected',
      selfConsistent: true,
    };
    writeReport(out, 'fixtures', [{ name: 'a', reference: 'pdf', slides: [sized] }], {
      poppler: '26.10.0',
      chromium: '156',
    });
    const summary = JSON.parse(readFileSync(join(out, 'summary.json'), 'utf8'));
    const slide = (summary.decks ?? summary)[0].slides[0];
    expect(slide).toMatchObject({
      sizeMismatch: true,
      renderSizePx: [1920, 1300],
      referenceSizePx: [1920, 1358],
      expectedRenderSizePx: [1920, 1358],
      expectedReferenceSizePx: [1920, 1358],
    });
    expect(slide.pixelMismatch).toBeUndefined();
    const html = readFileSync(join(out, 'index.html'), 'utf8');
    expect(html).toContain('pixels not compared (size differs)');
    expect(html).not.toContain('Probable cause: slide size differs');
    // A line cause of the same slide still shows.
    const both = { ...sized, cause: `${sized.cause}; line breaks differ` };
    writeReport(out, 'fixtures', [{ name: 'a', reference: 'pdf', slides: [both] }], {
      poppler: '26.10.0',
      chromium: '156',
    });
    expect(readFileSync(join(out, 'index.html'), 'utf8')).toContain(
      'Probable cause: line breaks differ<'
    );
  } finally {
    rmSync(out, { recursive: true, force: true });
  }
});

test('a baseline delta above the estimate offset is not blamed on the estimate', () => {
  const at = (baselineDelta: number): string =>
    probableCause({ ...failing, pixelMismatch: 0, baselineDelta }, []);
  expect(at(0.6)).toBe('line height or baseline model, or the PDF baseline estimate');
  // The bound has 2 decimals: 1.023 is the measured 1.02, and 1.03 is above it.
  expect(at(1.023)).toBe('line height or baseline model, or the PDF baseline estimate');
  expect(at(1.03)).toBe(
    'a baseline is off by more than the largest estimate offset seen (1.02 pt)'
  );
  expect(at(19.65)).toBe(
    'a baseline is off by more than the largest estimate offset seen (1.02 pt)'
  );
});

test('every fixed cause fits the 80 characters that pptx-eval keeps', () => {
  const lines = (over: Partial<SlideResult>): SlideResult => ({
    ...failing,
    lineMatch: 0.5,
    ...over,
  });
  for (const r of [
    lines({ unmatchedTexts: ['First'], pdfUnmatchedTexts: ['• First'] }),
    lines({ unmatchedTexts: ['a'], pdfUnmatchedTexts: [] }),
    lines({ unmatchedTexts: ['a b'], pdfUnmatchedTexts: ['a', 'b'], lineSplitOnly: true }),
    lines({ unmatchedTexts: ['a b'], pdfUnmatchedTexts: ['b a'] }),
    lines({ unmatchedTexts: ['a'], pdfUnmatchedTexts: ['c'] }),
    lines({ baselineDelta: 19.65 }),
  ])
    for (const cause of probableCause({ ...r, pixelMismatch: 0 }, []).split('; '))
      expect(cause.length).toBeLessThanOrEqual(80);
});
