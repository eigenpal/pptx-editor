import { describe, expect, test } from 'bun:test';
import { parseRunArgs } from './run-args.ts';

describe('the harness command line', () => {
  test('the fixture gate is the default; deck names and --update-gaps go with it', () => {
    expect(parseRunArgs([])).toEqual({
      mode: 'set',
      set: 'fixtures',
      decks: [],
      updateGaps: false,
      acceptWorse: false,
      lanes: ['editor', 'export'],
    });
    expect(parseRunArgs(['tables', 'colors', '--update-gaps'])).toEqual({
      mode: 'set',
      set: 'fixtures',
      decks: ['tables', 'colors'],
      updateGaps: true,
      acceptWorse: false,
      lanes: ['editor', 'export'],
    });
    expect(parseRunArgs(['tables', '--update-gaps', '--accept-worse'])).toMatchObject({
      updateGaps: true,
      acceptWorse: true,
    });
    expect(() => parseRunArgs(['--accept-worse'])).toThrow(
      '--accept-worse goes only with --update-gaps'
    );
    expect(parseRunArgs(['--set', 'corpus', 'deck2'])).toMatchObject({
      set: 'corpus',
      decks: ['deck2'],
      lanes: ['editor'],
    });
    expect(parseRunArgs(['--lane', 'export'])).toMatchObject({ lanes: ['export'] });
    expect(() => parseRunArgs(['--lane', 'print'])).toThrow('--lane must be editor or export');
    expect(() => parseRunArgs(['--list', 'j', '--out', 'o', '--lane', 'export'])).toThrow(
      '--lane does not go with --list'
    );
    expect(parseRunArgs(['-h'])).toEqual({ mode: 'help' });
  });

  test('corpus mode needs --out and takes its own flags', () => {
    expect(parseRunArgs(['--list', 'j.jsonl', '--out', 'o', '--shard', '2/3', '--resume'])).toEqual(
      {
        mode: 'list',
        list: 'j.jsonl',
        out: 'o',
        shard: '2/3',
        resume: true,
        fresh: false,
        images: false,
      }
    );
    expect(parseRunArgs(['--list', 'j', '--out', 'o', '--fresh'])).toMatchObject({ fresh: true });
    expect(() => parseRunArgs(['--list', 'j.jsonl'])).toThrow('--list needs --out');
  });

  test('an unknown flag or a flag of the other mode is a usage error', () => {
    expect(() => parseRunArgs(['--deck', 'tables'])).toThrow();
    expect(() => parseRunArgs(['--set', 'all'])).toThrow('--set must be fixtures or corpus');
    expect(() => parseRunArgs(['--resume'])).toThrow('--resume goes only with --list');
    expect(() => parseRunArgs(['--list', 'j', '--out', 'o', '--resume', '--fresh'])).toThrow(
      '--resume and --fresh do not go together'
    );
    expect(() => parseRunArgs(['--list', 'j', '--out', 'o', 'tables'])).toThrow(
      'deck names do not go with --list'
    );
    expect(() => parseRunArgs(['--set', 'corpus', '--update-gaps'])).toThrow(
      '--update-gaps goes only with the fixtures'
    );
  });
});
