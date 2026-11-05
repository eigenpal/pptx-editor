/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// The collaboration coverage gate (AGENTS.md "Collaboration verification"; DOCX Editor's
// `canonical-primitive-journal-coverage.test.ts`). Every editing command has a two-peer
// convergence fixture, or a reasoned entry in `COLLABORATION_UNCOVERED` with a probe that proves
// it writes nothing. Every kind of text input has a fixture. Then every fixture runs through two
// editors on linked Y.Docs. CONTRIBUTING.md, "Add a collaboration fixture", gives the steps.

import { describe, expect, test } from 'bun:test';
import type { EditorCommandType } from '@pptx-editor.dev/core/editor';
import { EDITOR_COMMAND_TYPES, TEXT_INPUT_KINDS } from '../../../core/src/editor/command-types.ts';
import { COLLABORATION_UNCOVERED } from '../collaboration/coverage-contract.ts';
import {
  CONVERGENCE_FIXTURES,
  NO_WRITE_PROBES,
  TEXT_INPUT_FIXTURES,
} from './collaboration-coverage-fixtures.ts';
import { converge, convergeInput, writesNothing } from './convergence-harness.ts';

const FIXTURES_FILE = 'packages/pro/src/__tests__/collaboration-coverage-fixtures.ts';
const CONTRACT_FILE = 'packages/pro/src/collaboration/coverage-contract.ts';
const HOW_TO = 'CONTRIBUTING.md, "Add a collaboration fixture", gives the steps.';

const commands = new Set<string>(EDITOR_COMMAND_TYPES);
const fixtures = Object.keys(CONVERGENCE_FIXTURES) as EditorCommandType[];
const probes = Object.keys(NO_WRITE_PROBES) as EditorCommandType[];
const uncovered = [...COLLABORATION_UNCOVERED.keys()];

describe('collaboration coverage gate', () => {
  test('every editing command has a convergence fixture or an uncovered reason', () => {
    const missing = EDITOR_COMMAND_TYPES.filter(
      (type) => !Object.hasOwn(CONVERGENCE_FIXTURES, type) && !COLLABORATION_UNCOVERED.has(type)
    );
    expect(
      missing,
      `These commands have no two-peer convergence fixture. Add one to CONVERGENCE_FIXTURES in ` +
        `${FIXTURES_FILE}. Only a command that writes nothing to the deck may instead get a ` +
        `reason in COLLABORATION_UNCOVERED (${CONTRACT_FILE}) and a probe. ${HOW_TO}`
    ).toEqual([]);
  });

  test('every uncovered entry names a current command, has no fixture, and gives a reason', () => {
    expect(
      uncovered.filter((type) => !commands.has(type)),
      `These COLLABORATION_UNCOVERED entries name no current command. Remove them.`
    ).toEqual([]);
    expect(
      uncovered.filter((type) => Object.hasOwn(CONVERGENCE_FIXTURES, type)),
      `These commands have a fixture, so remove their COLLABORATION_UNCOVERED entry.`
    ).toEqual([]);
    expect(
      uncovered.filter((type) => COLLABORATION_UNCOVERED.get(type)!.trim().length <= 40),
      `These COLLABORATION_UNCOVERED reasons are shorter than one sentence. Say why the ` +
        'command needs no fixture.'
    ).toEqual([]);
  });

  test('every fixture names a current command', () => {
    expect(
      fixtures.filter((type) => !commands.has(type)),
      `These CONVERGENCE_FIXTURES entries name no current command. Remove them.`
    ).toEqual([]);
  });

  test('every kind of text input has a convergence fixture', () => {
    expect(
      TEXT_INPUT_KINDS.filter((kind) => !Object.hasOwn(TEXT_INPUT_FIXTURES, kind)),
      `These text input kinds have no fixture in TEXT_INPUT_FIXTURES (${FIXTURES_FILE}).`
    ).toEqual([]);
  });

  test('every uncovered command, and only those, has a no-write probe', () => {
    expect(
      uncovered.filter((type) => !Object.hasOwn(NO_WRITE_PROBES, type)),
      `These COLLABORATION_UNCOVERED commands have no probe in NO_WRITE_PROBES ` +
        `(${FIXTURES_FILE}) to prove that they write nothing. ${HOW_TO}`
    ).toEqual([]);
    expect(
      probes.filter((type) => !COLLABORATION_UNCOVERED.has(type)),
      `These NO_WRITE_PROBES entries are for commands that are not in COLLABORATION_UNCOVERED.`
    ).toEqual([]);
  });
});

describe('two-peer convergence', () => {
  for (const type of fixtures)
    test(`${type} converges and survives save and reopen`, async () => {
      await converge(type, CONVERGENCE_FIXTURES[type]!);
    });
  for (const kind of TEXT_INPUT_KINDS)
    test(`text input ${kind} converges and survives save and reopen`, async () => {
      await convergeInput(kind, TEXT_INPUT_FIXTURES[kind]);
    });
});

describe('uncovered commands that write nothing', () => {
  for (const type of probes)
    test(`${type} runs and changes neither replica`, async () => {
      await writesNothing(type, NO_WRITE_PROBES[type]!);
    });
});
