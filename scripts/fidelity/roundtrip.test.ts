import { afterAll, describe, expect, setDefaultTimeout, test } from 'bun:test';
import {
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildPptx, para, run, shape } from '../fixtures/builder.ts';
import { ROUNDTRIP_SCHEMA } from './roundtrip-census.ts';
import { gitState, progressLine } from './roundtrip.ts';

// Each test starts the command several times; a loaded machine needs more than 5 s.
setDefaultTimeout(60_000);

const SCRIPT = join(import.meta.dir, 'roundtrip.ts');
const dir = mkdtempSync(join(tmpdir(), 'roundtrip-mode-'));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const deck = join(dir, 'deck.pptx');
writeFileSync(
  deck,
  buildPptx({
    slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 9, cy: 9, text: para(run('x')) }) }],
  })
);

function tool(...args: string[]): { code: number; stderr: string } {
  const r = Bun.spawnSync(['bun', SCRIPT, ...args], { stderr: 'pipe', stdout: 'pipe' });
  return { code: r.exitCode, stderr: r.stderr.toString() };
}

const lines = (out: string): Record<string, unknown>[] =>
  readFileSync(join(out, 'results.jsonl'), 'utf8')
    .trim()
    .split('\n')
    .map((line) => JSON.parse(line) as Record<string, unknown>);

describe('round-trip mode command', () => {
  test('usage errors print one line and exit 2; --help exits 0', () => {
    const list = join(dir, 'one.jsonl');
    writeFileSync(list, `${JSON.stringify({ id: 'one', deck, pdf: null })}\n`);
    expect(tool('--help').code).toBe(0);
    expect(tool('-h').code).toBe(0);
    for (const args of [
      ['--list', list, '--out', join(dir, 'o'), '--sharde', '1/2'],
      ['--list', list, '--out', '--resume'],
      ['--list', list, '--out', join(dir, 'o'), '--shard', '3/2'],
      ['--list', list, '--out', join(dir, 'o'), '--shard='],
      ['--list', list, '--out', join(dir, 'o'), '--saved='],
      ['--list', list, '--out', list],
      ['--list', join(dir, 'missing.jsonl'), '--out', join(dir, 'o')],
      ['--out', join(dir, 'o')],
    ]) {
      const r = tool(...args);
      expect(r.code).toBe(2);
      expect(r.stderr).toStartWith('roundtrip: ');
      expect(r.stderr).not.toContain('    at ');
      expect(r.stderr.trim().split('\n')).toHaveLength(1);
    }
    expect(existsSync(join(dir, 'o'))).toBe(false);
  });

  test('a run records every deck, keeps the copies, and resumes an interrupted deck once', () => {
    const list = join(dir, 'jobs.jsonl');
    const jobs = [
      { id: 'deck', deck, pdf: null },
      { id: 'gone', deck: join(dir, 'gone.pptx'), pdf: null },
      { id: 'hung', deck, pdf: null },
      { id: 'folder', deck: dir, pdf: null },
    ];
    writeFileSync(list, jobs.map((j) => JSON.stringify(j)).join('\n') + '\n');
    const out = join(dir, 'run');
    const saved = join(dir, 'saved');
    expect(tool('--list', list, '--out', out, '--saved', saved).code).toBe(0);
    // Pretend the process died while `hung` was in its edit step: the resume records it once,
    // with the facts of its finished pass 1.
    const hung = lines(out).find((r) => r.id === 'hung')!;
    const facts = {
      error: null,
      slides: hung.slides,
      parts: hung.parts,
      roundTrip: hung.roundTrip,
    };
    writeFileSync(join(out, 'current.json'), JSON.stringify({ id: 'hung', stage: 'edit', facts }));
    const results = lines(out).filter((r) => r.id !== 'hung');
    writeFileSync(
      join(out, 'results.jsonl'),
      results.map((r) => JSON.stringify(r)).join('\n') + '\n'
    );
    const resumed = Bun.spawnSync(
      ['bun', SCRIPT, '--list', list, '--out', out, '--saved', saved, '--resume'],
      { stdout: 'pipe', stderr: 'pipe' }
    );
    expect(resumed.exitCode).toBe(0);
    // The interrupted deck is reported, and counted as this invocation's pass-2 error.
    expect(resumed.stdout.toString()).toContain('hung: edit interrupted');
    const byId = Object.fromEntries(lines(out).map((r) => [r.id, r]));
    expect(byId.deck.error).toBeNull();
    expect(byId.gone.error).toEqual({ stage: 'open', reason: 'missing' });
    expect(byId.folder.error).toEqual({ stage: 'open', reason: 'unreadable' });
    expect(byId.hung.error).toBeNull();
    expect(byId.hung.roundTrip).toEqual(hung.roundTrip);
    expect((byId.hung.edit as { error: unknown }).error).toEqual({
      stage: 'edit',
      reason: 'interrupted',
    });
    expect(lines(out)).toHaveLength(4);
    expect(readdirSync(saved).sort()).toEqual(['deck.edited.pptx', 'deck.saved.pptx']);
    const runInfo = JSON.parse(readFileSync(join(out, 'run.json'), 'utf8'));
    expect(runInfo).toMatchObject({ schema: ROUNDTRIP_SCHEMA, keepsCopies: true });
    expect(runInfo.decks).toMatchObject({ total: 4, skipped: 3, finished: 1, editErrors: 1 });
    expect(runInfo.edit.recipe).toBeString();
  });

  test('a copy that cannot be written stops the run with exit 3, and the deck runs again', () => {
    const list = join(dir, 'keep.jsonl');
    writeFileSync(list, `${JSON.stringify({ id: 'k', deck, pdf: null })}\n`);
    const saved = join(dir, 'locked');
    mkdirSync(saved);
    chmodSync(saved, 0o500);
    try {
      const r = tool('--list', list, '--out', join(dir, 'keep'), '--saved', saved);
      expect(r.code).toBe(3);
      expect(r.stderr).toContain('cannot write a kept copy');
      expect(existsSync(join(dir, 'keep', 'current.json'))).toBe(false);
    } finally {
      chmodSync(saved, 0o700);
    }
    expect(
      tool('--list', list, '--out', join(dir, 'keep'), '--saved', saved, '--resume').code
    ).toBe(0);
    expect(lines(join(dir, 'keep'))[0].error).toBeNull();
  });

  test('a run, progress, or results file that cannot be written stops the run with exit 3', () => {
    const list = join(dir, 'files.jsonl');
    const two = [
      { id: 'f1', deck, pdf: null },
      { id: 'f2', deck, pdf: null },
    ];
    writeFileSync(list, two.map((j) => JSON.stringify(j)).join('\n') + '\n');
    const out = join(dir, 'files');
    const results = join(out, 'results.jsonl');
    const resume = () => tool('--list', list, '--out', out, '--resume');
    // Drop the last result in place, so that one deck runs again.
    const dropLast = () =>
      writeFileSync(results, readFileSync(results, 'utf8').split('\n')[0] + '\n');
    expect(tool('--list', list, '--out', out).code).toBe(0);
    for (const [locked, file] of [
      // A read-only folder stops the atomic run.json write first; the old run.json stays whole.
      [out, 'run.json'],
      [results, 'results.jsonl'],
    ] as const) {
      dropLast();
      chmodSync(locked, locked === out ? 0o500 : 0o400);
      try {
        const r = resume();
        expect(r.code).toBe(3);
        expect(r.stderr).toContain(`cannot write ${file}`);
      } finally {
        chmodSync(locked, locked === out ? 0o700 : 0o600);
      }
      expect(existsSync(join(out, 'current.json'))).toBe(false);
      expect(JSON.parse(readFileSync(join(out, 'run.json'), 'utf8')).schema).toBe(ROUNDTRIP_SCHEMA);
      expect(resume().code).toBe(0);
      expect(lines(out).map((line) => line.id)).toEqual(['f1', 'f2']);
    }
  });

  test('without --resume or --fresh, earlier results stop the run', () => {
    const list = join(dir, 'fresh.jsonl');
    writeFileSync(list, `${JSON.stringify({ id: 'r', deck, pdf: null })}\n`);
    const out = join(dir, 'fresh');
    expect(tool('--list', list, '--out', out).code).toBe(0);
    const again = tool('--list', list, '--out', out);
    expect(again.code).toBe(2);
    expect(again.stderr).toContain('--fresh to replace it');
    expect(lines(out)).toHaveLength(1);
    const startedAt = (): string =>
      JSON.parse(readFileSync(join(out, 'run.json'), 'utf8')).startedAt as string;
    const first = startedAt();
    expect(tool('--list', list, '--out', out, '--fresh').code).toBe(0);
    expect(lines(out)).toHaveLength(1);
    // A fresh run is a new run: it does not keep the start of the run it replaced.
    expect(startedAt()).not.toBe(first);
    // A results file that cannot be replaced is a fault of the machine, not a crash.
    chmodSync(out, 0o500);
    chmodSync(join(out, 'results.jsonl'), 0o400);
    try {
      const locked = tool('--list', list, '--out', out, '--fresh');
      expect(locked.code).toBe(3);
    } finally {
      chmodSync(out, 0o700);
      chmodSync(join(out, 'results.jsonl'), 0o600);
    }
    expect(tool('--list', list, '--out', out, '--fresh', '--resume').code).toBe(2);
    // An empty results file holds no results: a run without a flag goes on.
    const empty = join(dir, 'fresh-empty');
    mkdirSync(empty);
    writeFileSync(join(empty, 'results.jsonl'), '');
    expect(tool('--list', list, '--out', empty).code).toBe(0);
    expect(lines(empty)).toHaveLength(1);
  });

  test('an --out folder that cannot be made is a fault of the machine', () => {
    const list = join(dir, 'nofolder.jsonl');
    writeFileSync(list, `${JSON.stringify({ id: 'n', deck, pdf: null })}\n`);
    const parent = join(dir, 'locked-parent');
    mkdirSync(parent);
    chmodSync(parent, 0o500);
    try {
      const r = tool('--list', list, '--out', join(parent, 'out'));
      expect(r.code).toBe(3);
      expect(r.stderr).toContain('cannot write the --out folder');
    } finally {
      chmodSync(parent, 0o700);
    }
  });

  test('a resume without run.json names the cause, and one without git checks the sources', () => {
    const list = join(dir, 'cause.jsonl');
    writeFileSync(list, `${JSON.stringify({ id: 'c', deck, pdf: null })}\n`);
    const out = join(dir, 'cause');
    mkdirSync(out);
    writeFileSync(join(out, 'current.json'), JSON.stringify({ id: 'c', stage: 'open' }));
    const missing = tool('--list', list, '--out', out, '--resume');
    expect(missing.code).toBe(2);
    expect(missing.stderr).toContain('no readable run.json');
    // Results with no readable run.json (a write cut by a full disk) cannot prove their run.
    rmSync(join(out, 'current.json'));
    writeFileSync(join(out, 'results.jsonl'), `${JSON.stringify({ id: 'x', error: null })}\n`);
    writeFileSync(join(out, 'run.json'), '');
    const cut = tool('--list', list, '--out', out, '--resume');
    expect(cut.code).toBe(2);
    expect(cut.stderr).toContain('no readable run.json');
    rmSync(join(out, 'run.json'));
    expect(tool('--list', list, '--out', out, '--resume').code).toBe(2);
    // Without git the sources hash still checks the run: a new run, then a resume, both work,
    // and a resume of a run that recorded a commit names the commit as the difference.
    const noGit = (...args: string[]) =>
      Bun.spawnSync(['bun', SCRIPT, '--list', list, ...args], {
        env: { ...process.env, GIT_DIR: join(dir, 'no-git') },
        stderr: 'pipe',
      });
    const bare = join(dir, 'cause-no-git');
    expect(noGit('--out', bare).exitCode).toBe(0);
    expect(noGit('--out', bare, '--resume').exitCode).toBe(0);
    expect(JSON.parse(readFileSync(join(bare, 'run.json'), 'utf8')).sources).toMatch(
      /^[0-9a-f]{64}$/
    );
    const withGit = join(dir, 'cause-git');
    expect(tool('--list', list, '--out', withGit).code).toBe(0);
    const r = noGit('--out', withGit, '--resume');
    expect(r.exitCode).toBe(2);
    expect(r.stderr.toString()).toContain('output folder: commit');
  });

  test('a resume refuses a folder of another job list, shard, or engine', () => {
    const out = join(dir, 'other');
    const list = join(dir, 'a.jsonl');
    const two = [
      { id: 'a', deck, pdf: null },
      { id: 'b', deck, pdf: null },
    ];
    writeFileSync(list, two.map((j) => JSON.stringify(j)).join('\n') + '\n');
    expect(tool('--list', list, '--out', out, '--shard', '1/2').code).toBe(0);
    expect(tool('--list', list, '--out', out, '--shard', '1/2', '--resume').code).toBe(0);
    const shard = tool('--list', list, '--out', out, '--shard', '2/2', '--resume');
    expect(shard.code).toBe(2);
    expect(shard.stderr).toContain('output folder: --shard');
    const copies = tool(
      '--list',
      list,
      '--out',
      out,
      '--shard',
      '1/2',
      '--saved',
      join(dir, 's2'),
      '--resume'
    );
    expect(copies.stderr).toContain('output folder: --saved');
    const runFile = join(out, 'run.json');
    const runInfo = JSON.parse(readFileSync(runFile, 'utf8'));
    writeFileSync(runFile, JSON.stringify({ ...runInfo, commit: '0'.repeat(40) }));
    expect(tool('--list', list, '--out', out, '--shard', '1/2', '--resume').stderr).toContain(
      'output folder: commit'
    );
    writeFileSync(list, `${JSON.stringify(two[0])}\n`);
    writeFileSync(runFile, JSON.stringify(runInfo));
    const r = tool('--list', list, '--out', out, '--shard', '1/2', '--resume');
    expect(r.code).toBe(2);
    expect(r.stderr).toContain('output folder: job list');
  });
});

describe('engine identity', () => {
  test('only an edit in an engine path, tracked or untracked, changes the identity', () => {
    const repo = join(dir, 'repo');
    // A clean git environment: no GIT_* of the caller, and no signing of the test commit.
    const env = Object.fromEntries(
      Object.entries(process.env).filter(([key]) => !key.startsWith('GIT_'))
    );
    const git = (...args: string[]): void => {
      const r = Bun.spawnSync(['git', '-C', repo, '-c', 'commit.gpgsign=false', ...args], {
        stderr: 'pipe',
        env,
      });
      if (r.exitCode !== 0) throw new Error(r.stderr.toString());
    };
    for (const path of ['packages/core', 'docs', 'scripts/fidelity'])
      mkdirSync(join(repo, path), { recursive: true });
    writeFileSync(join(repo, 'packages/core/a.ts'), 'one');
    writeFileSync(join(repo, 'docs/a.md'), 'one');
    git('init', '-q');
    git('add', '.');
    git('-c', 'user.name=t', '-c', 'user.email=t@t', 'commit', '-q', '-m', 'one');
    const cwd = join(repo, 'scripts/fidelity');
    expect(gitState(cwd)).toMatchObject({ dirty: false, changes: null });
    writeFileSync(join(repo, 'docs/a.md'), 'two');
    expect(gitState(cwd)).toMatchObject({ dirty: false, changes: null });
    writeFileSync(join(repo, 'packages/core/a.ts'), 'two');
    const tracked = gitState(cwd);
    expect(tracked.dirty).toBe(true);
    writeFileSync(join(repo, 'scripts/fidelity/new.ts'), 'new');
    const untracked = gitState(cwd);
    expect(untracked.dirty).toBe(true);
    expect(untracked.changes).not.toBe(tracked.changes);
    expect(untracked.commit).toBe(tracked.commit);
    // An unreadable untracked file (here a broken link) still counts, by its error.
    symlinkSync(join(repo, 'missing'), join(repo, 'packages/core/broken.ts'));
    const broken = gitState(cwd);
    expect(broken.commit).toBe(tracked.commit);
    expect(broken.changes).not.toBe(untracked.changes);
  });
});

describe('progress line', () => {
  test('it names every check of both passes, or the error', () => {
    const items = { count: 0, items: [] };
    const line = progressLine({
      error: null,
      slides: 1,
      parts: null,
      roundTrip: {
        fingerprintEqual: true,
        xmlParts: 1,
        binaryParts: 0,
        digestStable: true,
        differences: items,
        namespaceDifferences: items,
      },
      edit: {
        edits: [
          { edit: 'setText', result: 'no-target' },
          { edit: 'move', result: 'applied' },
        ],
        error: null,
        editedParts: 1,
        fingerprintEqual: true,
        differences: items,
        collateral: { count: 2, items: [] },
        lost: 0,
        undoRestores: false,
      },
    });
    expect(line).toBe(
      'differences 0, namespace 0, digest stable, bytes 0/1; setText no-target, move applied, kept yes, collateral 2, lost 0, undo no'
    );
    expect(
      progressLine({
        error: { stage: 'digest', reason: 'exception', detail: 'TypeError' },
        slides: 1,
        parts: null,
        roundTrip: {
          fingerprintEqual: true,
          xmlParts: 1,
          binaryParts: 0,
          digestStable: null,
          differences: items,
          namespaceDifferences: items,
        },
        edit: {
          edits: [{ edit: 'setText', result: 'applied' }],
          error: null,
          editedParts: 1,
          fingerprintEqual: false,
          differences: { count: 1, items: [] },
          collateral: items,
          lost: 0,
          undoRestores: true,
        },
      })
    ).toBe(
      'digest exception; differences 0, namespace 0, digest -, bytes 0/1; setText applied, kept no 1, collateral 0, lost 0, undo ok'
    );
    expect(
      progressLine({
        error: { stage: 'open', reason: 'missing' },
        slides: null,
        parts: null,
        roundTrip: null,
        edit: null,
      })
    ).toBe('open missing');
  });
});
