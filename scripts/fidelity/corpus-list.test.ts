import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  CORPUS_SCHEMA,
  finishedIds,
  openResults,
  sourcesHash,
  startCorpusRun,
  writeJsonAtomic,
  parseJobList,
  parseShard,
  redactPaths,
  corpusThresholds,
  resultLine,
  shardJobs,
} from './corpus-list.ts';
import {
  compareLines,
  comparePixels,
  isListMarker,
  missingFonts,
  sizeMismatchCause,
  sizeMismatchDetail,
} from './compare.ts';
import { PNG } from 'pngjs';
import { THRESHOLDS } from './thresholds.ts';

describe('parseJobList', () => {
  test('reads JSON Lines and a JSON array', () => {
    const jobs = [
      { id: 'a1', deck: '/d/a.pptx', pdf: '/r/a.pdf' },
      { id: 'b2', deck: '/d/b.pptx', pdf: null },
    ];
    expect(parseJobList(jobs.map((j) => JSON.stringify(j)).join('\n'))).toEqual(jobs);
    expect(parseJobList(JSON.stringify(jobs))).toEqual(jobs);
    expect(parseJobList('{"id":"c","deck":"/d/c.pptx"}')).toEqual([
      { id: 'c', deck: '/d/c.pptx', pdf: null },
    ]);
  });

  test('rejects ids that could name another folder, duplicates, and relative paths', () => {
    expect(() => parseJobList('{"id":"../x","deck":"/d.pptx"}')).toThrow('invalid id');
    expect(() => parseJobList('{"id":"a/b","deck":"/d.pptx"}')).toThrow('invalid id');
    expect(() => parseJobList('{"id":".hidden","deck":"/d.pptx"}')).toThrow('invalid id');
    expect(() => parseJobList('{"id":"a","deck":"/d.pptx"}\n{"id":"a","deck":"/e.pptx"}')).toThrow(
      'duplicate id'
    );
    expect(() => parseJobList('{"id":"a","deck":"d.pptx"}')).toThrow('absolute');
    expect(() => parseJobList('{"id":"a","deck":"/d.pptx","pdf":"r.pdf"}')).toThrow('absolute');
  });
});

describe('shards', () => {
  test('parseShard accepts i/n and rejects the rest', () => {
    expect(parseShard('2/3')).toEqual({ index: 2, count: 3 });
    for (const bad of ['0/3', '4/3', '1/0', 'x', '1-2']) expect(() => parseShard(bad)).toThrow();
  });

  test('every job is in exactly one shard', () => {
    const jobs = Array.from({ length: 10 }, (_, i) => i);
    const shards = [1, 2, 3].map((index) => shardJobs(jobs, { index, count: 3 }));
    expect(shards.flat().sort((a, b) => a - b)).toEqual(jobs);
    expect(shards[0]).toEqual([0, 3, 6, 9]);
  });
});

describe('finishedIds', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'pptx-corpus-'));
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  test('reads the ids of complete lines and skips a line cut short', () => {
    const file = join(dir, 'results.jsonl');
    expect(finishedIds(file).size).toBe(0);
    writeFileSync(file, '{"id":"a1","slides":[]}\n{"id":"b2","slides":[]}\n{"id":"c3","sli');
    expect([...finishedIds(file)]).toEqual(['a1', 'b2']);
  });

  test('a resume cuts off a line that a crash left short, so the next line stays readable', () => {
    const file = join(dir, 'results.jsonl');
    writeFileSync(file, '{"id":"a1","slides":[]}\n{"id":"b2","sli');
    expect([...openResults(file, true)]).toEqual(['a1']);
    writeFileSync(file, `${readFileSync(file, 'utf8')}{"id":"b2","slides":[]}\n`, 'utf8');
    expect([...finishedIds(file)]).toEqual(['a1', 'b2']);
  });

  test('a run without resume starts with no results', () => {
    const file = join(dir, 'results.jsonl');
    writeFileSync(file, '{"id":"a1","slides":[]}\n');
    expect(openResults(file, false).size).toBe(0);
    expect(finishedIds(file).size).toBe(0);
  });
});

describe('no paths in the corpus output', () => {
  test('redactPaths removes paths and addresses, and everything after a path', () => {
    expect(redactPaths('Command failed: pdfinfo -f 1 /Users/me/Client deck.pdf\nstderr')).toBe(
      'Command failed: pdfinfo -f 1 <path>'
    );
    expect(redactPaths('page.goto: failed at http://127.0.0.1:5555/harness.html')).toBe(
      'page.goto: failed at <address>'
    );
    expect(redactPaths("can't read '~/decks/x.pptx'")).toBe("can't read '<path>");
    expect(redactPaths('save/reopen differs; 3/4 lines')).toBe('save/reopen differs; 3/4 lines');
  });

  test('run.json records every threshold, with the page size tolerance', () => {
    expect(corpusThresholds()).toMatchObject({ ...THRESHOLDS, pageSizeTolerancePt: 0.625 });
  });

  test('a size mismatch has both sizes and no pixel score in its result line', () => {
    const line = resultLine(
      { id: 'abc', deck: '/d.pptx', pdf: '/d.pdf' },
      {
        name: 'abc',
        reference: 'pdf',
        slides: [
          {
            index: 1,
            status: 'fail',
            diagnostics: [],
            cause: 'slide size differs: the render does not have its expected size',
            sizeMismatch: true,
            renderSizePx: [1920, 1300],
            referenceSizePx: [1920, 1358],
          },
        ],
      }
    );
    expect(JSON.parse(line).slides[0]).toMatchObject({
      sizeMismatch: true,
      renderSizePx: [1920, 1300],
      referenceSizePx: [1920, 1358],
      pixelMismatch: null,
    });
  });

  test('a result line keeps scores and drops every path', () => {
    const line = resultLine(
      { id: 'abc', deck: '/secret/deck.pptx', pdf: '/secret/deck.pdf' },
      {
        name: 'abc',
        error: 'the harness failed: at /secret/deck.pptx',
        reference: 'stale',
        referenceIssues: ['the PDF does not open (pdfinfo failed) /secret/deck.pdf'],
        slides: [
          {
            index: 1,
            status: 'stale',
            diagnostics: ['font substituted: Calibri'],
            cause: 'the PDF does not open (pdfinfo failed) /secret/deck.pdf',
          },
        ],
      }
    );
    expect(line).not.toContain('secret');
    const withText = resultLine(
      { id: 'abc', deck: '/d.pptx', pdf: null },
      {
        name: 'abc',
        slides: [
          {
            index: 1,
            status: 'fail',
            diagnostics: [],
            cause: '',
            unmatchedTexts: ['Client quarterly revenue'],
            pdfUnmatchedTexts: ['Client quarterly revenue'],
            fontsSubstituted: { reference: ['Aptos'], editor: [] },
            diffBoxPt: [1, 2, 3, 4],
          },
        ],
      }
    );
    expect(withText).not.toContain('Client');
    expect(JSON.parse(withText).slides[0]).toMatchObject({
      fontsSubstituted: { reference: ['Aptos'], editor: [] },
      diffBoxPt: [1, 2, 3, 4],
    });
    const parsed = JSON.parse(line);
    expect(parsed).toMatchObject({
      id: 'abc',
      reference: 'stale',
      error: 'the harness failed: at <path>',
    });
    expect(parsed.slides[0]).toMatchObject({
      lineMatch: null,
      baselineDeltaPt: null,
      pdfLinesUnmatched: null,
      estimatedBaselines: null,
      skippedLines: null,
    });
  });

  test('a result line has the schema 2 line counts', () => {
    const line = resultLine(
      { id: 'abc', deck: '/d.pptx', pdf: '/d.pdf' },
      {
        name: 'abc',
        slides: [
          {
            index: 1,
            status: 'pass',
            diagnostics: [],
            cause: '',
            estimatedBaselines: 1,
            skippedLines: 2,
          },
        ],
      }
    );
    expect(JSON.parse(line).slides[0]).toMatchObject({ estimatedBaselines: 1, skippedLines: 2 });
  });
});

describe('compareLines', () => {
  const line = (text: string, baselinePt: number) => ({ node: 'n', text, baselinePt, leftPt: 0 });
  const pdf = (text: string, yMin: number, yMax: number) => ({
    text,
    xMin: 0,
    xMax: 100,
    yMin,
    yMax,
  });

  test('matches lines by normalized text and measures the estimated baseline', () => {
    const score = compareLines(
      [line('Hello  world', 98), line('Second', 130), line('   ', 0)],
      [pdf('Hello world', 80, 100), pdf('Other', 120, 140)]
    );
    expect(score.matched).toBe(1);
    expect(score.total).toBe(2);
    expect(score.unmatched).toEqual(['Second']);
    expect(score.worstBaseline).toBeCloseTo(2, 6); // 100 − 0.2 × 20 = 96
  });

  test('ignores whitespace that pdftotext adds, and joins pieces on one baseline', () => {
    const score = compareLines(
      [line('Spaced out tight', 98), line('small LARGE line', 130)],
      [
        pdf('S p a c e d o u t tight', 80, 100),
        { ...pdf('small', 115, 132.5), xMin: 0, baseline: 130 },
        { ...pdf('LARGE line', 110, 135), xMin: 30, baseline: 130.12 },
      ]
    );
    expect(score.matched).toBe(2);
    expect(score.unmatched).toEqual([]);
  });

  test('a bullet line matches the PDF text that starts with the bullet', () => {
    const score = compareLines(
      [
        { ...line('Item', 98), bullet: '•' },
        { ...line('Next', 130), bullet: '1.', numbered: true },
        { ...line('Deep', 162), bullet: '(iii)', numbered: true },
        { ...line('Symbol', 194), bullet: '§' },
      ],
      [
        pdf('• Item', 80, 100),
        pdf('1. Next', 112, 132.5),
        pdf('(iii) Deep', 144, 164.5),
        pdf('\uf0a7 Symbol', 176, 196.5),
      ]
    );
    expect(score.matched).toBe(4);
  });

  test('a line off the slide is skipped only when the PDF does not hold it', () => {
    const score = compareLines(
      [
        { ...line('Above the page', -2), offSlide: true },
        { ...line('Also above', -20), offSlide: true },
      ],
      [pdf('Also above', -38, -17.5)]
    );
    expect(score.skipped).toBe(1);
    expect(score.matched).toBe(1);
    expect(score.total).toBe(1);
  });

  test('a char bullet accepts one more character than itself, not a missing word', () => {
    const score = compareLines(
      [{ ...line('world', 98), bullet: '•' }],
      [pdf('• the world', 80, 100)]
    );
    expect(score.matched).toBe(0);
  });

  test('a line off the slide with a different line break still fails', () => {
    const score = compareLines(
      [{ ...line('Hello there', -2), offSlide: true }],
      [pdf('Hello', -16, 2)]
    );
    expect(score.skipped).toBe(0);
    expect(score.unmatched).toEqual(['Hello there']);
  });

  test('a bulleted or numbered line off the slide with a different line break still fails', () => {
    const score = compareLines(
      [
        { ...line('Hello there', -2), offSlide: true, bullet: '•' },
        { ...line('Next step', -40), offSlide: true, bullet: '1.', numbered: true },
        { ...line('Gone', -80), offSlide: true, bullet: '2.', numbered: true },
      ],
      [pdf('• Hello', -16, 2), pdf('1. Next', -54, -36), pdf('3. Gone', -94, -76)]
    );
    expect(score.skipped).toBe(1);
    expect(score.unmatched).toEqual(['Hello there', 'Next step']);
  });

  test('a wrong number or a piece across a gap does not match', () => {
    const score = compareLines(
      [{ ...line('Next', 130), bullet: '1.', numbered: true }, line('Total cost', 162)],
      [
        pdf('2. Next', 112, 132.5),
        { ...pdf('Total', 150, 164.5), xMax: 30, baseline: 162 },
        { ...pdf('cost', 150, 164.5), xMin: 300, xMax: 330, baseline: 162 },
      ]
    );
    expect(score.matched).toBe(0);
  });

  test('pairs repeated text with the nearest PDF line, and reads upside-down boxes from the top', () => {
    const score = compareLines(
      [line('Label', 300), line('Label', 100), { ...line('Label', 204), upsideDown: true }],
      [pdf('Label', 80, 100), pdf('Label', 200, 220), pdf('Label', 280, 300)]
    );
    expect(score.matched).toBe(3);
    // 296 vs 300, 96 vs 100, and the upside-down box 200–220 has its baseline at 204.
    expect(score.worstBaseline).toBeCloseTo(4, 6);
  });

  test('counts the PDF lines that no laid-out line matched', () => {
    const score = compareLines(
      [line('Hello', 98)],
      [pdf('Hello', 80, 100), pdf('Rotated label', 0, 10), pdf('  ', 0, 10)]
    );
    expect(score.pdfUnmatched).toBe(1);
  });

  test('a line that the PDF splits into words on one row is a split, not a break', () => {
    const at = (text: string, xMin: number, yMin: number) => ({
      text,
      xMin,
      xMax: xMin + 6 * text.length,
      yMin,
      yMax: yMin + 20,
    });
    // `spacing`: two font sizes on one row; `text-runs`: letter spacing in one PDF line.
    const sizes = compareLines(
      [line('small LARGE line height', 98)],
      [at('LARGE line height', 300, 80), at('small', 100, 82)]
    );
    expect(sizes.splitOnly).toBe(true);
    const spaced = compareLines(
      [line('Spaced out tight', 98)],
      [at('S p a c e d o u t tight', 0, 80)]
    );
    // Letter spacing inside one PDF line matches: the line check ignores whitespace (`holds`).
    expect(spaced.matched).toBe(1);
    // `autofit`: a word moved to the next row is a real line break.
    const moved = compareLines(
      [line('The quick brown', 98), line('fox jumps', 128)],
      [at('The quick', 0, 80), at('brown fox jumps', 0, 110)]
    );
    expect(moved.splitOnly).toBe(false);
    // A long list line: the flags read the full texts, not the 60 characters of the report.
    const long = 'A list line that is longer than the sixty characters the report keeps';
    const listed = compareLines([line(long, 98)], [at(`• ${long}`, 0, 80)]);
    expect(listed.markersOnly).toBe(true);
    // A lost word is not a marker: the editor dropped text.
    const lost = compareLines([line('quick brown fox', 98)], [at('The quick brown fox', 0, 80)]);
    expect(lost.markersOnly).toBe(false);
    expect(['•', '1.', 'a)', '(iv)', '10.'].every(isListMarker)).toBe(true);
    expect(['The', 'ab.', '1', '12345.', '“', '(', '$', 'civil.'].some(isListMarker)).toBe(false);
    expect(['\uF0A7', '▪', '–'].every(isListMarker)).toBe(true);
    expect(listed.unmatched[0].length).toBe(60);
    const broken = compareLines(
      [line(`${long} here`, 98), line('and more', 128)],
      [at(long, 0, 80), at('here and more', 0, 110)]
    );
    expect([broken.sameText, broken.splitOnly]).toEqual([true, false]);
    // The same pieces on two rows are two lines, not one split line.
    const rows = compareLines([line('small LARGE', 98)], [at('small', 0, 80), at('LARGE', 0, 110)]);
    expect(rows.splitOnly).toBe(false);
  });
});

describe('starting a corpus run', () => {
  let out: string;
  beforeEach(() => {
    out = mkdtempSync(join(tmpdir(), 'pptx-run-'));
  });
  afterEach(() => rmSync(out, { recursive: true, force: true }));

  const identity = {
    schema: 1,
    list: 'l1',
    shard: { index: 1, count: 2 },
    sources: 's1',
    thresholds: { maxPixelMismatch: 0.005 },
    popplerVersion: '26.10.0',
    referenceFonts: true,
    fonts: 'f1',
  };
  const results = (): void =>
    writeFileSync(join(out, 'results.jsonl'), '{"id":"a1","slides":[]}\n');
  const run = (over: object = {}): void =>
    writeFileSync(join(out, 'run.json'), JSON.stringify({ ...identity, ...over, decks: {} }));

  test('a new folder starts with nothing done', () => {
    expect(startCorpusRun(out, identity, { resume: false, fresh: false })).toEqual({
      done: new Set(),
    });
  });

  test('existing results stop a run without --resume, and --fresh replaces them', () => {
    results();
    expect(startCorpusRun(out, identity, { resume: false, fresh: false })).toEqual({
      refused:
        'results.jsonl exists in --out. Add --resume to continue that run, or --fresh to replace it.',
    });
    expect(startCorpusRun(out, identity, { resume: false, fresh: true })).toEqual({
      done: new Set(),
    });
    expect(finishedIds(join(out, 'results.jsonl')).size).toBe(0);
  });

  test('a resume of the same run skips the finished ids', () => {
    results();
    run();
    expect(startCorpusRun(out, identity, { resume: true, fresh: false })).toEqual({
      done: new Set(['a1']),
    });
  });

  test('a resume refuses results whose run.json is missing', () => {
    results();
    expect(startCorpusRun(out, identity, { resume: true, fresh: false })).toMatchObject({
      refused: expect.stringContaining('run.json of the earlier run'),
    });
  });

  test('a resume refuses an earlier run with other fonts', () => {
    results();
    run({ fonts: 'f0' });
    expect(startCorpusRun(out, identity, { resume: true, fresh: false })).toMatchObject({
      refused: expect.stringContaining('these differ from the earlier run: fonts'),
    });
  });

  test('a resume refuses a run of another format', () => {
    results();
    run({ schema: 1 });
    expect(
      startCorpusRun(out, { ...identity, schema: CORPUS_SCHEMA }, { resume: true, fresh: false })
    ).toMatchObject({
      refused: expect.stringContaining('these differ from the earlier run: schema'),
    });
  });

  test('a resume refuses an earlier run with other sources or another shard', () => {
    results();
    run({ sources: 's0', shard: { index: 2, count: 2 } });
    expect(startCorpusRun(out, identity, { resume: true, fresh: false })).toEqual({
      refused:
        '--resume needs the same run, but these differ from the earlier run: shard, sources. Use a new --out folder, or --fresh to replace the results.',
    });
  });
});

describe('sourcesHash', () => {
  let root: string;
  beforeEach(() => {
    root = mkdtempSync(join(tmpdir(), 'pptx-src-'));
    mkdirSync(join(root, 'src/__tests__'), { recursive: true });
    writeFileSync(join(root, 'src/a.ts'), 'a');
    writeFileSync(join(root, 'src/a.test.ts'), 't');
    writeFileSync(join(root, 'src/__tests__/b.ts'), 'b');
  });
  afterEach(() => rmSync(root, { recursive: true, force: true }));

  test('changes with a source edit, a new file, or a missing path, not with a test', () => {
    const first = sourcesHash(root, ['src']);
    writeFileSync(join(root, 'src/a.test.ts'), 'changed');
    writeFileSync(join(root, 'src/__tests__/b.ts'), 'changed');
    expect(sourcesHash(root, ['src'])).toBe(first);
    writeFileSync(join(root, 'src/a.ts'), 'edited');
    const edited = sourcesHash(root, ['src']);
    expect(edited).not.toBe(first);
    writeFileSync(join(root, 'src/new.ts'), 'n');
    expect(sourcesHash(root, ['src'])).not.toBe(edited);
    expect(sourcesHash(root, ['src', 'gone'])).not.toBe(sourcesHash(root, ['src']));
  });
});

describe('missingFonts', () => {
  test('a family matches the PostScript names that start with it', () => {
    const present = ['Calibri', 'Calibri-Bold', 'ArialMT', 'TimesNewRomanPSMT', 'SegoeUI'];
    expect(missingFonts(['Calibri', 'Arial', 'Times New Roman', 'Segoe UI'], present)).toEqual([]);
    expect(missingFonts(['Aptos', 'Calibri', 'Aptos'], present)).toEqual(['Aptos']);
  });
});

describe('the difference box', () => {
  const image = (fill: (x: number, y: number) => number): Buffer => {
    const p = new PNG({ width: 20, height: 10 });
    for (let y = 0; y < 10; y += 1)
      for (let x = 0; x < 20; x += 1) {
        const i = (y * 20 + x) * 4;
        const v = fill(x, y);
        p.data[i] = p.data[i + 1] = p.data[i + 2] = v;
        p.data[i + 3] = 255;
      }
    return PNG.sync.write(p);
  };

  test('is the box around the pixels that differ, in pixels', () => {
    const white = image(() => 255);
    const block = image((x, y) => (x >= 5 && x < 9 && y >= 2 && y < 4 ? 0 : 255));
    expect(comparePixels(white, block).diffBox).toEqual([5, 2, 9, 4]);
    expect(comparePixels(white, white).diffBox).toBeUndefined();
  });
});

describe('the size check', () => {
  const blank = (width: number, height: number): Buffer => {
    const p = new PNG({ width, height });
    p.data.fill(255);
    return PNG.sync.write(p);
  };
  const size = (width: number, height: number) => ({ width, height });

  test('without expected sizes, the two images must have the same size', () => {
    expect(comparePixels(blank(20, 10), blank(20, 10)).sizeMismatch).toBe(false);
    expect(comparePixels(blank(20, 10), blank(17, 10)).sizeMismatch).toBe(true);
  });

  test('a reference that is short where its page ends compares the overlap', () => {
    const expected = { render: size(20, 10), reference: size(17, 10) };
    expect(comparePixels(blank(20, 10), blank(17, 10), { expected })).toMatchObject({
      sizeMismatch: false,
      mismatch: 0,
    });
  });

  test('a render of the wrong size fails, even when the reference has its expected size', () => {
    const expected = { render: size(20, 10), reference: size(20, 10) };
    const r = comparePixels(blank(20, 7), blank(20, 10), { expected });
    expect(r).toMatchObject({ sizeMismatch: true, mismatch: 1 });
    expect(sizeMismatchCause(r.sizes)).toBe(
      'slide size differs: the render does not have its expected size'
    );
    expect(sizeMismatchDetail(r.sizes)).toBe(
      'the render is 20 × 7 px, 20 × 10 px expected, and the PDF raster is 20 × 10 px, as expected'
    );
    // pptx-eval keeps 80 characters of a cause.
    const both = { ...r.sizes, reference: size(1, 1) };
    for (const c of [sizeMismatchCause(r.sizes), sizeMismatchCause(both)])
      expect(c.length).toBeLessThanOrEqual(80);
    // `eigenpal/pptx-eval` splits causes on ";", and groups equal causes.
    expect(sizeMismatchCause(r.sizes)).not.toContain(';');
    expect(sizeMismatchDetail(r.sizes)).not.toContain(';');
  });

  test('every score has both sizes, also when they match', () => {
    const expected = { render: size(20, 10), reference: size(17, 10) };
    expect(comparePixels(blank(20, 10), blank(17, 10), { expected }).sizes).toEqual({
      render: size(20, 10),
      reference: size(17, 10),
      expected,
    });
  });

  test('a reference of the wrong size fails', () => {
    const expected = { render: size(20, 10), reference: size(17, 10) };
    expect(comparePixels(blank(20, 10), blank(20, 10), { expected }).sizeMismatch).toBe(true);
  });
});

describe('writeJsonAtomic', () => {
  test('writes the whole file and leaves no temporary file', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-json-'));
    try {
      const file = join(dir, 'run.json');
      writeJsonAtomic(file, { a: 1 });
      writeJsonAtomic(file, { a: 2 });
      expect(JSON.parse(readFileSync(file, 'utf8'))).toEqual({ a: 2 });
      expect(readdirSync(dir)).toEqual(['run.json']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
