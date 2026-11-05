import { describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../contracts/index.ts';
import { EDITOR_COMMAND_TYPES, TEXT_INPUT_KINDS } from '../command-types.ts';
import { isPictureCommand } from '../picture-commands.ts';
import { isSlideCommand } from '../slide-command-table.ts';
import { isTextFormatCommand } from '../text-format-commands.ts';

// The run-time list of command types. The compiler checks it against `EditorCommand`; this
// test checks that it holds no duplicate and that the sub-tables that route commands to their
// compilers name only listed commands.

const probe = (type: string) => ({ type }) as unknown as EditorCommand;

describe('EDITOR_COMMAND_TYPES', () => {
  test('lists each command once', () => {
    expect(new Set(EDITOR_COMMAND_TYPES).size).toBe(EDITOR_COMMAND_TYPES.length);
    expect(EDITOR_COMMAND_TYPES.length).toBeGreaterThan(100);
  });

  test('lists each text input kind once', () => {
    expect(new Set(TEXT_INPUT_KINDS).size).toBe(TEXT_INPUT_KINDS.length);
    expect(TEXT_INPUT_KINDS).toContain('text');
  });

  test('routes each listed command to one compiler sub-table at most', () => {
    const routes = [isPictureCommand, isSlideCommand, isTextFormatCommand];
    for (const route of routes) expect(route(probe('notACommand'))).toBe(false);
    for (const type of EDITOR_COMMAND_TYPES)
      expect(routes.filter((route) => route(probe(type))).length).toBeLessThanOrEqual(1);
    expect(EDITOR_COMMAND_TYPES.some((type) => isPictureCommand(probe(type)))).toBe(true);
    expect(EDITOR_COMMAND_TYPES.some((type) => isSlideCommand(probe(type)))).toBe(true);
    expect(EDITOR_COMMAND_TYPES.some((type) => isTextFormatCommand(probe(type)))).toBe(true);
  });
});
