import { describe, expect, test } from 'bun:test';
import { parseRunArgs } from './run-args.ts';

describe('the harness command line', () => {
  test('the fixture gate is the default; --deck and --print-gaps go with it', () => {
    expect(parseRunArgs([])).toEqual({
      mode: 'set',
      set: 'fixtures',
      deck: undefined,
      printGaps: false,
    });
    expect(parseRunArgs(['--set', 'fixtures', '--deck=tables', '--print-gaps'])).toEqual({
      mode: 'set',
      set: 'fixtures',
      deck: 'tables',
      printGaps: true,
    });
    expect(parseRunArgs(['--set', 'corpus', '--deck', 'deck2'])).toMatchObject({ set: 'corpus' });
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

  test('an unknown flag, a stray word, or a flag of the other mode is a usage error', () => {
    expect(() => parseRunArgs(['--dek', 'tables'])).toThrow();
    expect(() => parseRunArgs(['tables'])).toThrow();
    expect(() => parseRunArgs(['--deck'])).toThrow();
    expect(() => parseRunArgs(['--set', 'all'])).toThrow('--set must be fixtures or corpus');
    expect(() => parseRunArgs(['--resume'])).toThrow('--resume goes only with --list');
    expect(() => parseRunArgs(['--list', 'j', '--out', 'o', '--resume', '--fresh'])).toThrow(
      '--resume and --fresh do not go together'
    );
    expect(() => parseRunArgs(['--list', 'j', '--out', 'o', '--deck', 'x'])).toThrow(
      '--deck does not go with --list'
    );
    expect(() => parseRunArgs(['--set', 'corpus', '--print-gaps'])).toThrow(
      '--print-gaps goes only with the fixtures'
    );
  });
});
