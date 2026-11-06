import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  approximationTitles,
  otherDiagnostics,
  probableCause,
  writeReport,
  type SlideResult,
} from './report.ts';

const failing: SlideResult = {
  index: 1,
  status: 'fail',
  pixelMismatch: 0.02,
  lineMatch: 1,
  baselineDelta: 0.1,
  diagnostics: [],
  cause: '',
};

describe('probable causes', () => {
  test('an unrendered effect and the approximations a slide uses are causes', () => {
    expect(probableCause(failing, ['effect not rendered: presetShdw'])).toBe(
      'effect not rendered (presetShdw)'
    );
    const d = ['approximation: blur', 'approximation: glow-soft-edge'];
    expect(probableCause(failing, d)).toBe('approximations: Blur, Glow and soft edge');
    expect(approximationTitles(d)).toEqual(['Blur', 'Glow and soft edge']);
  });

  test('a slide with no known cause points to the diff', () => {
    expect(probableCause(failing, [])).toBe(
      'pixels differ, and no layout diagnostic names a cause: see the diff'
    );
  });

  test('the line model is a cause only when a line check fails', () => {
    const d = ['approximation: line-model'];
    expect(probableCause(failing, d)).toBe(
      'pixels differ, and no layout diagnostic names a cause: see the diff'
    );
    expect(probableCause({ ...failing, baselineDelta: 0.6 }, d)).toBe(
      'line height or baseline model; approximations: Line model'
    );
  });

  test('line causes say whether the text breaks elsewhere or is missing', () => {
    const lines = { ...failing, pixelMismatch: 0, lineMatch: 0.5 };
    expect(probableCause({ ...lines, missingLines: ['a b'], extraLines: ['a', 'b'] }, [])).toBe(
      'line breaks differ'
    );
    expect(probableCause({ ...lines, missingLines: ['a'], extraLines: [] }, [])).toBe(
      'our page lacks text of the reference page'
    );
    expect(probableCause({ ...lines, missingLines: ['a'], extraLines: ['c'] }, [])).toBe(
      'line breaks or text differ'
    );
  });

  test('approximations and other diagnostics are listed apart', () => {
    const d = ['font substituted: Aptos', 'approximation: line-model'];
    expect(otherDiagnostics(d)).toEqual(['font substituted: Aptos']);
    expect(approximationTitles(d)).toEqual(['Line model']);
  });
});

test('the report shows both pages and the diff of each slide, and links both PDFs', () => {
  const out = mkdtempSync(join(tmpdir(), 'pptx-report-'));
  try {
    writeReport(
      out,
      'fixtures',
      [
        {
          name: 'deck',
          reference: 'pdf',
          oursPdf: join(out, 'deck', 'ours.pdf'),
          referencePdf: join(out, 'deck', 'reference.pdf'),
          slides: [
            {
              ...failing,
              ours: join(out, 'deck', 'slide1-ours.png'),
              reference: join(out, 'deck', 'slide1-reference.png'),
              diff: join(out, 'deck', 'slide1-diff.png'),
              cause: 'line breaks differ',
            },
          ],
        },
      ],
      { tools: 'pdf.js 5', gate: { problems: ['deck/1: a new failure'], next: 'Fix it.' } }
    );
    const html = readFileSync(join(out, 'index.html'), 'utf8');
    for (const part of [
      'deck/ours.pdf',
      'deck/reference.pdf',
      'deck/slide1-ours.png',
      'deck/slide1-reference.png',
      'deck/slide1-diff.png',
      'Gate: fails',
      'deck/1: a new failure',
      'Probable cause: line breaks differ',
    ])
      expect(html).toContain(part);
    const summary = JSON.parse(readFileSync(join(out, 'summary.json'), 'utf8'));
    // The summary keeps scores, not image paths.
    expect(summary[0].slides[0]).toMatchObject({ index: 1, pixelMismatch: 0.02 });
    expect(summary[0].slides[0].ours).toBeUndefined();
  } finally {
    rmSync(out, { recursive: true, force: true });
  }
});
