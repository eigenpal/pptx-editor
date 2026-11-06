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
            missingLines: ['Client quarterly revenue'],
            extraLines: ['Client quarterly revenue'],
            fontsSubstituted: ['Aptos'],
            diffBoxPt: [1, 2, 3, 4],
          },
        ],
      }
    );
    expect(withText).not.toContain('Client');
    expect(JSON.parse(withText).slides[0]).toMatchObject({
      fontsSubstituted: ['Aptos'],
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
    });
    expect(CORPUS_SCHEMA).toBe(3);
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
    pdfjs: '5.4.624',
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
