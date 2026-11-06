/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Fixed histories of the text fuzzer (`yjs-text-fuzz.test.ts`): the 32 runs of the random text
// test that it replaces, seeds 1 to 16 in its two modes, recorded as actions. Each one keeps
// that run's client ids and edits: in each of 4 rounds every peer makes one or two edits with
// the links held, the links deliver to the peers in a random order, and a flush ends the round.
// The fuzzer replays them with its own checks, which are stricter than that test's.
//
// One string per round, one action per comma-separated token:
// - `p1 type 0:3 x` types "x" at offset 3 of paragraph 0 on peer 1;
// - `p1 enter 0:3+2!` and `p1 type 0:3+2! x` replace the 2 characters from offset 3;
// - `p1 delete 0:3+2` deletes them; `p1 bold 0:3+2` and `p1 italic 0:3+0` format them (an
//   empty selection formats the caret);
// - `p1 undo`, `p1 redo`;
// - `>p2` delivers every update that waits for peer 2, in the order they were sent;
// - `flush` delivers every update, and the round's checks run.

import type { FuzzKnownFailure, FuzzRegression } from './fuzz-support.ts';
import type { EditKind, TextAction } from './yjs-text-fuzz.test.ts';

const TOKEN = /^p(\d+) (\w+)(?: (\d+):(\d+)(?:\+(\d+))?(!)?)?(?: (\S))?$/u;

/** One round's actions from its compact form. */
function decode(round: string): TextAction[] {
  return round.split(', ').map((token): TextAction => {
    if (token === 'flush') return { kind: 'flush' };
    if (token.startsWith('>p')) return { kind: 'deliverTo', to: Number(token.slice(2)) };
    const m = TOKEN.exec(token);
    if (!m) throw new Error(`not an action: ${JSON.stringify(token)}`);
    const [, peer, edit, paragraph, offset, length, over, letter] = m;
    return {
      kind: 'edit',
      peer: Number(peer),
      edit: edit as EditKind,
      paragraph: Number(paragraph ?? 0),
      offset: Number(offset ?? 0),
      length: Number(length ?? 0),
      ...(over ? { over: true as const } : {}),
      ...(letter ? { letter } : {}),
    };
  });
}

const historyOf =
  (what: string) =>
  (seed: number, clients: number[], rounds: string[]): FuzzRegression<TextAction> => ({
    name: `random history ${seed}: ${what}`,
    seed,
    clients,
    actions: rounds.flatMap(decode),
  });

/** A history that fails today under the stricter checks, with its issue and its problem. */
const failing = (
  regression: FuzzRegression<TextAction>,
  issue: string,
  problem: string
): FuzzKnownFailure<TextAction> => ({ ...regression, issue, problem });

const textHistory = historyOf('typing, Enter, and delete');
const formatHistory = historyOf('with formats, undo, and redo');
const undoHistory = historyOf('typing, Enter, delete, undo, and redo over 6 rounds');
const longFormatHistory = historyOf('with formats, undo, and redo over 6 rounds');

/** Typing, Enter, and delete: replayed by the `TEXT` mode. */
export const TEXT_HISTORIES: readonly FuzzRegression<TextAction>[] = [
  textHistory(
    1,
    [1001, 3003, 2002],
    [
      'p0 enter 1:1+2!, p0 type 1:1 Ā, p1 type 0:2 ā, p2 type 0:6 Ă, >p1, >p2, >p0, flush',
      'p0 enter 0:2+2!, p0 type 3:0 ă, p1 type 2:0 Ą, p1 type 0:4 ą, p2 type 1:0 Ć, p2 delete 1:2+1, >p2, >p0, >p1, flush',
      'p0 delete 3:2+1, p0 type 3:2 ć, p1 delete 1:5+1, p1 delete 2:0+2, p2 type 3:0 Ĉ, >p1, >p2, >p0, flush',
      'p0 type 2:0 ĉ, p0 delete 1:0+1, p1 type 0:2 Ċ, p2 enter 0:0+1!, >p0, >p2, >p1, flush',
    ]
  ),
  textHistory(
    2,
    [3003, 1001, 2002],
    [
      'p0 type 1:3 Ā, p0 type 1:0 ā, p1 enter 1:2, p2 type 1:0 Ă, >p2, >p1, >p0, flush',
      'p0 type 1:3 ă, p0 enter 1:2, p1 type 0:8 Ą, p2 type 0:8 ą, p2 delete 2:0+3, >p0, >p2, >p1, flush',
      'p0 enter 2:3, p1 enter 0:9+1!, p1 type 1:0 Ć, p2 type 0:10 ć, >p0, >p1, >p2, flush',
      'p0 enter 2:0+2!, p0 type 4:3 Ĉ, p1 delete 3:0+1, p2 type 4:0 ĉ, >p0, >p2, >p1, flush',
    ]
  ),
  textHistory(
    3,
    [3003, 1001, 2002],
    [
      'p0 type 1:2 Ā, p1 enter 1:0, p2 enter 1:4, >p1, >p0, >p2, flush',
      'p0 type 3:0 ā, p0 type 1:0 Ă, p1 type 0:3 ă, p1 enter 0:4+3!, p2 type 3:0 Ą, p2 type 1:0 ą, >p1, >p2, >p0, flush',
      'p0 enter 4:1+1!, p0 delete 4:0+1, p1 type 3:5 Ć, p1 type 0:2 ć, p2 enter 1:1+1!, p2 type 2:0 Ĉ, >p0, >p1, >p2, flush',
      'p0 enter 0:2, p0 type 6:0 ĉ, p1 delete 4:5+1, p1 type 2:1 Ċ, p2 type 0:3 ċ, p2 type 6:0 Č, >p1, >p0, >p2, flush',
    ]
  ),
  textHistory(
    4,
    [3003, 1001, 2002],
    [
      'p0 delete 0:3+1, p1 enter 1:3+1!, p1 type 2:0 Ā, p2 type 1:4 ā, >p1, >p0, >p2, flush',
      'p0 type 2:1 Ă, p1 type 1:2 ă, p2 type 1:1 Ą, p2 type 1:1 ą, >p0, >p1, >p2, flush',
      'p0 type 1:1 Ć, p0 type 0:6 ć, p1 delete 2:1+1, p2 enter 2:1+1!, >p1, >p2, >p0, flush',
      'p0 enter 3:0, p0 enter 4:0, p1 type 1:8 Ĉ, p1 type 3:0 ĉ, p2 type 0:0 Ċ, >p0, >p2, >p1, flush',
    ]
  ),
  textHistory(
    5,
    [1001, 2002, 3003],
    [
      'p0 enter 1:0+2!, p1 enter 1:1+3!, p2 delete 0:4+1, p2 type 0:6 Ā, >p1, >p0, >p2, flush',
      'p0 type 1:0 ā, p1 type 0:7 Ă, p2 type 1:0 ă, p2 type 2:0 Ą, >p0, >p2, >p1, flush',
      'p0 type 1:2 ą, p0 type 3:0 Ć, p1 delete 0:8+1, p2 enter 1:2, >p1, >p2, >p0, flush',
      'p0 enter 1:3, p0 delete 5:0+1, p1 enter 2:0, p1 type 0:2 ć, p2 type 3:0 Ĉ, >p0, >p1, >p2, flush',
    ]
  ),
  textHistory(
    6,
    [1001, 3003, 2002],
    [
      'p0 type 1:0 Ā, p1 type 0:8 ā, p1 enter 0:4+3!, p2 delete 0:1+3, >p1, >p0, >p2, flush',
      'p0 delete 2:4+1, p0 type 2:1 Ă, p1 enter 0:1, p2 enter 2:3+2!, p2 delete 2:1+2, >p2, >p1, >p0, flush',
      'p0 type 0:0 ă, p0 type 1:0 Ą, p1 type 0:1 ą, p2 enter 2:0+1!, >p1, >p2, >p0, flush',
      'p0 enter 0:2+1!, p1 enter 0:2+1!, p1 type 1:0 Ć, p2 enter 3:0+1!, >p0, >p1, >p2, flush',
    ]
  ),
  textHistory(
    7,
    [2002, 3003, 1001],
    [
      'p0 type 1:2 Ā, p0 type 1:4 ā, p1 type 1:0 Ă, p1 type 1:1 ă, p2 type 0:1 Ą, p2 delete 1:2+1, >p0, >p1, >p2, flush',
      'p0 delete 1:4+3, p0 type 0:1 ą, p1 type 0:7 Ć, p2 type 1:5 ć, >p2, >p1, >p0, flush',
      'p0 type 0:2 Ĉ, p0 enter 0:8+3!, p1 enter 1:5, p1 delete 1:4+1, p2 delete 1:4+1, p2 enter 0:6+1!, >p0, >p1, >p2, flush',
      'p0 enter 2:0, p0 type 4:2 ĉ, p1 enter 0:7, p2 delete 3:2+2, >p1, >p0, >p2, flush',
    ]
  ),
  textHistory(
    8,
    [2002, 1001, 3003],
    [
      'p0 delete 1:2+1, p1 type 1:0 Ā, p2 enter 0:1+3!, p2 enter 0:0, >p1, >p0, >p2, flush',
      'p0 enter 0:0, p0 type 1:0 ā, p1 enter 2:4, p2 type 2:1 Ă, p2 type 1:1 ă, >p1, >p2, >p0, flush',
      'p0 enter 1:1, p0 type 3:0 Ą, p1 type 5:4 ą, p2 type 3:4 Ć, p2 enter 0:0, >p1, >p0, >p2, flush',
      'p0 delete 2:0+1, p1 type 5:6 ć, p1 delete 4:1+1, p2 enter 4:1+2!, >p1, >p0, >p2, flush',
    ]
  ),
  textHistory(
    9,
    [2002, 1001, 3003],
    [
      'p0 type 1:3 Ā, p1 enter 0:5+2!, p1 enter 0:5, p2 delete 0:4+3, p2 enter 0:5, >p1, >p2, >p0, flush',
      'p0 type 2:0 ā, p1 type 0:0 Ă, p1 type 3:0 ă, p2 enter 0:2+1!, p2 enter 4:0, >p2, >p0, >p1, flush',
      'p0 enter 4:1, p0 enter 7:5, p1 enter 4:1, p2 type 5:0 Ą, p2 delete 0:2+1, >p1, >p0, >p2, flush',
      'p0 enter 7:1, p1 type 3:1 ą, p1 enter 1:1, p2 enter 8:0, p2 delete 7:0+1, >p2, >p0, >p1, flush',
    ]
  ),
  textHistory(
    10,
    [1001, 2002, 3003],
    [
      'p0 type 1:2 Ā, p0 delete 0:0+3, p1 delete 0:3+1, p2 type 1:4 ā, >p0, >p1, >p2, flush',
      'p0 enter 0:1+1!, p1 type 1:0 Ă, p2 type 0:0 ă, >p1, >p0, >p2, flush',
      'p0 type 0:1 Ą, p0 enter 2:5, p1 enter 1:0+2!, p1 type 2:0 ą, p2 type 2:0 Ć, >p1, >p0, >p2, flush',
      'p0 type 2:0 ć, p1 type 4:0 Ĉ, p1 type 4:3 ĉ, p2 type 1:0 Ċ, >p2, >p0, >p1, flush',
    ]
  ),
  textHistory(
    11,
    [1001, 2002, 3003],
    [
      'p0 type 1:4 Ā, p0 type 1:3 ā, p1 type 0:6 Ă, p1 type 0:2 ă, p2 type 0:5 Ą, p2 type 1:4 ą, >p2, >p0, >p1, flush',
      'p0 enter 1:2, p0 delete 1:1+1, p1 type 1:6 Ć, p2 type 1:3 ć, >p0, >p1, >p2, flush',
      'p0 enter 1:0+1!, p1 type 1:0 Ĉ, p1 type 2:3 ĉ, p2 type 1:1 Ċ, p2 type 1:1 ċ, >p1, >p0, >p2, flush',
      'p0 enter 2:0, p0 type 2:0 Č, p1 type 1:1 č, p1 type 2:0 Ď, p2 enter 2:0, p2 type 3:0 ď, >p0, >p1, >p2, flush',
    ]
  ),
  textHistory(
    12,
    [2002, 3003, 1001],
    [
      'p0 enter 0:0+2!, p0 delete 1:4+1, p1 enter 1:3+1!, p2 delete 1:1+1, p2 type 0:3 Ā, >p2, >p1, >p0, flush',
      'p0 enter 1:1+2!, p1 type 2:1 ā, p2 enter 0:0, p2 type 0:0 Ă, >p2, >p1, >p0, flush',
      'p0 type 0:1 ă, p0 type 2:0 Ą, p1 type 3:2 ą, p2 delete 0:0+1, >p0, >p2, >p1, flush',
      'p0 type 1:0 Ć, p0 type 0:1 ć, p1 enter 1:0, p1 type 2:0 Ĉ, p2 type 5:0 ĉ, p2 type 3:1 Ċ, >p1, >p2, >p0, flush',
    ]
  ),
  textHistory(
    13,
    [3003, 1001, 2002],
    [
      'p0 type 0:0 Ā, p1 type 1:3 ā, p1 enter 0:5+1!, p2 enter 0:5+1!, p2 type 1:0 Ă, >p2, >p1, >p0, flush',
      'p0 enter 1:3, p1 type 1:1 ă, p2 type 0:1 Ą, p2 type 0:2 ą, >p1, >p0, >p2, flush',
      'p0 delete 0:2+3, p1 delete 0:2+1, p1 enter 0:1, p2 enter 4:2+1!, p2 enter 5:1+1!, >p0, >p1, >p2, flush',
      'p0 enter 4:0, p0 type 1:1 Ć, p1 delete 0:0+1, p2 type 2:4 ć, p2 type 2:0 Ĉ, >p1, >p2, >p0, flush',
    ]
  ),
  textHistory(
    14,
    [2002, 3003, 1001],
    [
      'p0 type 0:3 Ā, p1 delete 1:0+1, p1 type 0:1 ā, p2 delete 0:5+1, p2 delete 0:6+1, >p2, >p1, >p0, flush',
      'p0 type 0:1 Ă, p0 delete 0:8+1, p1 enter 1:1+1!, p1 delete 0:3+2, p2 type 0:0 ă, >p2, >p1, >p0, flush',
      'p0 delete 1:0+1, p1 enter 1:1, p2 delete 1:0+1, >p2, >p0, >p1, flush',
      'p0 type 2:0 Ą, p0 delete 3:0+1, p1 enter 1:0, p2 enter 1:0, >p1, >p0, >p2, flush',
    ]
  ),
  textHistory(
    15,
    [3003, 2002, 1001],
    [
      'p0 type 0:5 Ā, p1 type 0:3 ā, p2 delete 0:1+1, p2 type 1:4 Ă, >p2, >p1, >p0, flush',
      'p0 type 1:0 ă, p1 enter 1:0+2!, p1 delete 0:0+1, p2 type 1:3 Ą, p2 enter 1:1, >p0, >p2, >p1, flush',
      'p0 type 1:0 ą, p0 type 1:0 Ć, p1 enter 0:8, p2 delete 2:1+3, >p0, >p1, >p2, flush',
      'p0 delete 4:0+1, p1 type 2:2 ć, p1 enter 1:0, p2 delete 0:3+2, p2 type 1:0 Ĉ, >p2, >p1, >p0, flush',
    ]
  ),
  textHistory(
    16,
    [1001, 3003, 2002],
    [
      'p0 enter 0:0+3!, p1 type 1:2 Ā, p2 type 1:4 ā, >p0, >p2, >p1, flush',
      'p0 enter 1:4+1!, p1 delete 2:0+1, p1 enter 0:0, p2 type 1:3 Ă, >p0, >p1, >p2, flush',
      'p0 delete 2:0+1, p0 type 4:0 ă, p1 type 3:0 Ą, p2 delete 2:1+2, >p2, >p1, >p0, flush',
      'p0 delete 3:0+1, p1 type 0:0 ą, p1 type 2:2 Ć, p2 enter 3:0+1!, >p2, >p0, >p1, flush',
    ]
  ),
];

/** Typing, Enter, delete, bold, italic, undo, and redo: replayed by the `FORMAT` mode. */
export const FORMAT_HISTORIES: readonly FuzzRegression<TextAction>[] = [
  formatHistory(
    1,
    [1001, 3003, 2002],
    [
      'p0 italic 1:1+2, p0 delete 0:7+1, p1 delete 0:1+1, p2 type 0:6 Ā, >p1, >p2, >p0, flush',
      'p0 bold 0:1+2, p0 delete 1:0+2, p1 type 1:1 ā, p1 type 0:3 Ă, p2 type 1:0 ă, p2 redo, >p2, >p0, >p1, flush',
      'p0 redo, p0 type 1:4 Ą, p1 redo, p1 redo, p2 type 1:0 ą, >p1, >p2, >p0, flush',
      'p0 undo, p0 redo, p1 enter 0:5+2!, p2 italic 0:0+1, >p0, >p2, >p1, flush',
    ]
  ),
  formatHistory(
    2,
    [3003, 1001, 2002],
    [
      'p0 enter 1:3+1!, p0 type 2:0 Ā, p1 italic 1:2+0, p2 type 1:0 ā, >p2, >p1, >p0, flush',
      'p0 type 1:3 Ă, p0 bold 1:2+0, p1 type 0:8 ă, p2 type 0:8 Ą, p2 redo, >p0, >p2, >p1, flush',
      'p0 italic 2:1+0, p1 bold 0:9+1, p1 delete 0:9+1, p2 delete 0:9+1, >p0, >p1, >p2, flush',
      'p0 bold 1:1+2, p0 enter 2:1, p1 undo, p2 delete 2:0+1, >p0, >p2, >p1, flush',
    ]
  ),
  formatHistory(
    3,
    [3003, 1001, 2002],
    [
      'p0 type 1:2 Ā, p1 italic 1:0+0, p2 italic 1:4+0, >p1, >p0, >p2, flush',
      'p0 redo, p0 delete 0:7+1, p1 type 0:3 ā, p1 bold 0:4+3, p2 type 1:3 Ă, p2 delete 0:4+3, >p1, >p2, >p0, flush',
      'p0 italic 1:4+1, p0 redo, p1 type 1:6 ă, p1 enter 0:3+2!, p2 bold 0:2+2, p2 redo, >p0, >p1, >p2, flush',
      'p0 italic 0:1+0, p0 redo, p1 undo, p1 delete 0:3+2, p2 enter 0:2+1!, p2 type 3:0 Ą, >p1, >p0, >p2, flush',
    ]
  ),
  formatHistory(
    4,
    [3003, 1001, 2002],
    [
      'p0 redo, p1 italic 1:3+1, p1 delete 1:2+2, p2 delete 1:3+1, >p1, >p0, >p2, flush',
      'p0 delete 1:1+1, p1 type 0:4 Ā, p2 type 0:2 ā, p2 type 1:0 Ă, >p0, >p1, >p2, flush',
      'p0 delete 0:2+2, p0 type 0:6 ă, p1 redo, p2 italic 1:1+1, >p1, >p2, >p0, flush',
      'p0 italic 1:0+1, p0 italic 1:1+1, p1 type 0:9 Ą, p1 redo, p2 type 0:0 ą, >p0, >p2, >p1, flush',
    ]
  ),
  formatHistory(
    5,
    [1001, 2002, 3003],
    [
      'p0 italic 1:0+2, p1 italic 1:1+3, p2 redo, p2 enter 0:6+2!, >p1, >p0, >p2, flush',
      'p0 redo, p1 delete 0:5+1, p2 undo, p2 undo, >p0, >p2, >p1, flush',
      'p0 type 0:6 Ā, p0 undo, p1 redo, p2 italic 0:6+1, >p1, >p2, >p0, flush',
      'p0 italic 0:7+0, p0 undo, p1 bold 1:4+0, p1 type 0:2 ā, p2 enter 1:0, >p0, >p1, >p2, flush',
    ]
  ),
  formatHistory(
    6,
    [1001, 3003, 2002],
    [
      'p0 redo, p1 type 0:8 Ā, p1 italic 0:4+3, p2 redo, >p1, >p0, >p2, flush',
      'p0 redo, p0 type 1:1 ā, p1 bold 0:8+1, p2 italic 1:2+2, p2 undo, >p2, >p1, >p0, flush',
      'p0 enter 0:3+1!, p0 type 0:1 Ă, p1 type 0:9 ă, p2 italic 0:0+1, >p1, >p2, >p0, flush',
      'p0 italic 0:2+2, p1 italic 0:2+1, p1 type 0:0 Ą, p2 bold 1:2+2, >p0, >p1, >p2, flush',
    ]
  ),
  formatHistory(
    7,
    [2002, 3003, 1001],
    [
      'p0 type 1:2 Ā, p0 type 1:4 ā, p1 enter 1:0+1!, p1 type 1:0 Ă, p2 enter 0:1, p2 redo, >p0, >p1, >p2, flush',
      'p0 undo, p0 delete 0:0+1, p1 delete 0:0+1, p2 enter 3:4+1!, >p2, >p1, >p0, flush',
      'p0 enter 1:1+2!, p0 bold 0:0+0, p1 bold 3:3+0, p1 undo, p2 undo, p2 italic 0:0+0, >p0, >p1, >p2, flush',
      'p0 italic 2:4+0, p0 type 3:1 ă, p1 italic 0:0+0, p2 redo, >p1, >p0, >p2, flush',
    ]
  ),
  formatHistory(
    8,
    [2002, 1001, 3003],
    [
      'p0 undo, p1 type 1:0 Ā, p2 italic 0:1+3, p2 bold 0:1+0, >p1, >p0, >p2, flush',
      'p0 bold 0:4+1, p0 undo, p1 italic 1:5+0, p2 type 1:1 ā, p2 delete 0:7+1, >p1, >p2, >p0, flush',
      'p0 bold 0:6+0, p0 type 1:1 Ă, p1 type 1:5 ă, p2 delete 1:5+1, p2 italic 0:3+1, >p1, >p0, >p2, flush',
      'p0 undo, p1 type 1:7 Ą, p1 redo, p2 bold 1:2+2, >p1, >p0, >p2, flush',
    ]
  ),
  formatHistory(
    9,
    [2002, 1001, 3003],
    [
      'p0 enter 1:3+1!, p1 bold 0:5+2, p1 italic 0:8+0, p2 undo, p2 italic 0:7+0, >p1, >p2, >p0, flush',
      'p0 undo, p1 enter 0:0+1!, p1 redo, p2 bold 0:4+1, p2 bold 2:0+0, >p2, >p0, >p1, flush',
      'p0 italic 1:5+2, p0 italic 2:4+0, p1 italic 1:5+2, p2 redo, p2 undo, >p1, >p0, >p2, flush',
      'p0 bold 2:4+0, p1 type 0:0 Ā, p1 bold 0:1+0, p2 bold 2:0+0, p2 redo, >p2, >p0, >p1, flush',
    ]
  ),
  formatHistory(
    10,
    [1001, 2002, 3003],
    [
      'p0 type 1:2 Ā, p0 redo, p1 redo, p2 delete 1:3+1, >p0, >p1, >p2, flush',
      'p0 bold 0:3+1, p1 type 1:0 ā, p2 delete 0:1+3, >p1, >p0, >p2, flush',
      'p0 type 0:2 Ă, p0 bold 1:4+0, p1 italic 1:0+2, p1 delete 1:4+1, p2 type 1:0 ă, >p1, >p0, >p2, flush',
      'p0 delete 0:1+3, p1 enter 1:0, p1 type 2:5 Ą, p2 undo, >p2, >p0, >p1, flush',
    ]
  ),
  formatHistory(
    11,
    [1001, 2002, 3003],
    [
      'p0 type 1:4 Ā, p0 delete 1:3+1, p1 type 0:6 ā, p1 delete 0:2+1, p2 delete 0:5+2, p2 enter 1:4, >p2, >p0, >p1, flush',
      'p0 italic 2:0+0, p0 undo, p1 type 2:0 Ă, p2 enter 2:0, >p0, >p1, >p2, flush',
      'p0 italic 2:0+1, p1 type 1:2 ă, p1 enter 2:0+1!, p2 type 1:4 Ą, p2 enter 2:1, >p1, >p0, >p2, flush',
      'p0 bold 3:0+0, p0 type 2:0 ą, p1 type 1:2 Ć, p1 type 4:0 ć, p2 bold 3:0+0, p2 type 3:0 Ĉ, >p0, >p1, >p2, flush',
    ]
  ),
  formatHistory(
    12,
    [2002, 3003, 1001],
    [
      'p0 italic 0:0+2, p0 redo, p1 bold 1:3+1, p2 redo, p2 type 0:3 Ā, >p2, >p1, >p0, flush',
      'p0 bold 0:1+2, p1 delete 1:0+1, p2 bold 0:0+2, p2 enter 0:1+2!, >p2, >p1, >p0, flush',
      'p0 type 0:2 ā, p0 type 1:0 Ă, p1 delete 1:2+1, p2 redo, >p0, >p2, >p1, flush',
      'p0 redo, p0 enter 0:4, p1 bold 0:2+1, p1 redo, p2 delete 2:2+1, p2 enter 1:1+3!, >p1, >p2, >p0, flush',
    ]
  ),
  formatHistory(
    13,
    [3003, 1001, 2002],
    [
      'p0 enter 0:0, p1 enter 1:3+1!, p1 bold 0:5+1, p2 italic 0:5+1, p2 delete 0:0+2, >p2, >p1, >p0, flush',
      'p0 italic 1:5+0, p1 type 1:1 Ā, p2 type 0:0 ā, p2 type 0:0 Ă, >p1, >p0, >p2, flush',
      'p0 redo, p1 undo, p1 italic 0:0+0, p2 italic 3:0+0, p2 italic 3:0+0, >p0, >p1, >p2, flush',
      'p0 italic 2:3+0, p0 delete 0:1+2, p1 redo, p2 type 1:5 ă, p2 enter 1:0+1!, >p1, >p2, >p0, flush',
    ]
  ),
  formatHistory(
    14,
    [2002, 3003, 1001],
    [
      'p0 type 0:3 Ā, p1 undo, p1 delete 0:1+1, p2 undo, p2 redo, >p2, >p1, >p0, flush',
      'p0 enter 0:1+2!, p0 undo, p1 bold 1:1+1, p1 undo, p2 type 0:0 ā, >p2, >p1, >p0, flush',
      'p0 redo, p1 italic 1:3+1, p2 redo, >p2, >p0, >p1, flush',
      'p0 type 1:5 Ă, p0 redo, p1 italic 1:5+0, p2 italic 0:1+1, >p1, >p0, >p2, flush',
    ]
  ),
  formatHistory(
    15,
    [3003, 2002, 1001],
    [
      'p0 enter 0:5+3!, p1 type 0:3 Ā, p2 redo, p2 type 1:4 ā, >p2, >p1, >p0, flush',
      'p0 enter 2:0+2!, p1 italic 1:0+0, p1 redo, p2 type 1:0 Ă, p2 italic 2:1+0, >p0, >p2, >p1, flush',
      'p0 redo, p0 delete 1:0+1, p1 bold 0:6+0, p2 undo, >p0, >p1, >p2, flush',
      'p0 undo, p1 type 1:0 ă, p1 italic 0:4+0, p2 undo, p2 type 0:0 Ą, >p2, >p1, >p0, flush',
    ]
  ),
];

/**
 * Histories that fail today under the stricter checks, each with its issue and the problem it
 * pins (see `FuzzKnownFailure`).
 */
export const FORMAT_HISTORY_FAILURES: readonly FuzzKnownFailure<TextAction>[] = [
  failing(
    formatHistory(
      16,
      [1001, 3003, 2002],
      [
        'p0 bold 0:0+3, p1 enter 1:2, p2 delete 1:3+1, >p0, >p2, >p1, flush',
        'p0 bold 1:2+0, p1 undo, p1 italic 0:8+0, p2 delete 1:1+1, >p0, >p1, >p2, flush',
        'p0 undo, p0 enter 1:0+1!, p1 type 1:0 Ā, p2 redo, >p2, >p1, >p0, flush',
        'p0 redo, p1 enter 0:4+3!, p1 type 2:0 ā, p2 italic 1:0+0, >p2, >p0, >p1, flush',
      ]
    ),
    '#3',
    'p1 undo added characters it may not add: "L"'
  ),
];

// Seeds of the removed random text test that once lost text, showed it twice, or diverged when
// it ran 6 rounds with as many undos and redos as edits, or with formats, undo, and redo. Each
// one is fixed; it stays as a history of its 6 rounds so the case stays fixed.

/** Typing, Enter, delete, undo, and redo over 6 rounds: replayed by the `TEXT` mode. */
export const UNDO_HISTORIES: readonly FuzzRegression<TextAction>[] = [
  undoHistory(
    7,
    [2002, 3003, 1001],
    [
      'p0 type 1:2 Ā, p0 type 1:4 ā, p1 enter 1:0+1!, p1 type 1:0 Ă, p2 enter 0:1, p2 redo, >p0, >p1, >p2, flush',
      'p0 redo, p0 delete 0:0+1, p1 delete 0:0+1, p2 enter 3:4+1!, >p2, >p1, >p0, flush',
      'p0 enter 1:1+2!, p0 undo, p1 undo, p1 redo, p2 redo, p2 undo, >p0, >p1, >p2, flush',
      'p0 undo, p0 type 2:1 ă, p1 undo, p2 redo, >p1, >p0, >p2, flush',
      'p0 type 4:0 Ą, p0 type 4:1 ą, p1 undo, p2 redo, p2 redo, >p0, >p2, >p1, flush',
      'p0 delete 2:0+1, p1 type 0:0 Ć, p2 type 4:2 ć, >p1, >p2, >p0, flush',
    ]
  ),
  undoHistory(
    112,
    [1001, 2002, 3003],
    [
      'p0 delete 0:1+2, p0 type 1:3 Ā, p1 delete 0:1+1, p2 type 1:0 ā, >p1, >p0, >p2, flush',
      'p0 redo, p1 redo, p1 undo, p2 redo, >p2, >p0, >p1, flush',
      'p0 undo, p0 undo, p1 undo, p2 enter 0:1, >p0, >p2, >p1, flush',
      'p0 delete 0:0+1, p0 undo, p1 undo, p2 type 2:1 Ă, >p1, >p0, >p2, flush',
      'p0 undo, p1 undo, p2 redo, >p2, >p0, >p1, flush',
      'p0 redo, p0 redo, p1 redo, p2 redo, p2 undo, >p1, >p0, >p2, flush',
    ]
  ),
  undoHistory(
    209,
    [2002, 1001, 3003],
    [
      'p0 type 0:3 Ā, p0 redo, p1 enter 1:1, p1 redo, p2 redo, >p0, >p1, >p2, flush',
      'p0 type 2:3 ā, p1 undo, p2 type 2:3 Ă, p2 delete 1:0+1, >p1, >p0, >p2, flush',
      'p0 undo, p0 redo, p1 redo, p1 enter 2:2+1!, p2 delete 2:1+1, p2 redo, >p0, >p2, >p1, flush',
      'p0 undo, p0 enter 1:0, p1 redo, p2 redo, >p1, >p0, >p2, flush',
      'p0 enter 0:8+1!, p0 enter 0:8, p1 delete 3:0+2, p1 enter 2:0, p2 redo, >p2, >p0, >p1, flush',
      'p0 enter 6:0, p1 redo, p2 type 4:0 ă, p2 type 7:1 Ą, >p1, >p2, >p0, flush',
    ]
  ),
  undoHistory(
    228,
    [1001, 2002, 3003],
    [
      'p0 enter 1:4, p1 enter 1:4, p2 undo, >p1, >p0, >p2, flush',
      'p0 type 3:0 Ā, p1 undo, p1 enter 0:5+1!, p2 undo, p2 undo, >p1, >p2, >p0, flush',
      'p0 enter 2:3, p0 enter 4:0, p1 redo, p1 undo, p2 redo, p2 type 0:2 ā, >p1, >p0, >p2, flush',
      'p0 type 5:0 Ă, p1 undo, p2 undo, p2 delete 2:0+1, >p1, >p0, >p2, flush',
      'p0 enter 2:0, p1 type 4:0 ă, p2 undo, p2 delete 4:0+1, >p1, >p2, >p0, flush',
      'p0 type 4:0 Ą, p1 redo, p1 redo, p2 redo, >p0, >p1, >p2, flush',
    ]
  ),
  undoHistory(
    304,
    [2002, 3003, 1001],
    [
      'p0 undo, p0 undo, p1 enter 0:1+2!, p2 undo, >p2, >p0, >p1, flush',
      'p0 delete 2:1+2, p0 type 1:1 Ā, p1 undo, p2 type 1:3 ā, p2 redo, >p0, >p2, >p1, flush',
      'p0 delete 1:0+1, p0 type 0:8 Ă, p1 undo, p1 undo, p2 redo, >p2, >p0, >p1, flush',
      'p0 redo, p1 undo, p1 type 1:2 ă, p2 type 0:9 Ą, >p2, >p1, >p0, flush',
      'p0 undo, p0 redo, p1 type 1:0 ą, p1 type 0:1 Ć, p2 redo, >p1, >p0, >p2, flush',
      'p0 redo, p1 type 1:4 ć, p1 type 1:0 Ĉ, p2 redo, >p2, >p1, >p0, flush',
    ]
  ),
  undoHistory(
    327,
    [2002, 1001, 3003],
    [
      'p0 type 1:4 Ā, p1 type 0:3 ā, p2 redo, >p0, >p1, >p2, flush',
      'p0 delete 0:6+1, p0 type 0:2 Ă, p1 undo, p1 delete 0:2+3, p2 type 1:1 ă, p2 undo, >p0, >p1, >p2, flush',
      'p0 type 1:4 Ą, p0 type 1:1 ą, p1 redo, p2 redo, p2 type 0:2 Ć, >p1, >p2, >p0, flush',
      'p0 undo, p1 undo, p2 redo, p2 type 1:5 ć, >p1, >p0, >p2, flush',
      'p0 enter 0:3, p1 type 0:4 Ĉ, p1 enter 0:9+1!, p2 redo, p2 enter 0:0+2!, >p0, >p1, >p2, flush',
      'p0 redo, p1 delete 2:5+1, p2 redo, p2 redo, >p0, >p1, >p2, flush',
    ]
  ),
  undoHistory(
    363,
    [2002, 1001, 3003],
    [
      'p0 redo, p1 undo, p1 delete 1:1+1, p2 undo, p2 delete 1:1+3, >p1, >p0, >p2, flush',
      'p0 undo, p0 type 1:1 Ā, p1 enter 1:0, p1 redo, p2 enter 1:1, >p1, >p0, >p2, flush',
      'p0 type 2:0 ā, p0 undo, p1 undo, p1 redo, p2 undo, p2 undo, >p1, >p2, >p0, flush',
      'p0 enter 2:3+2!, p0 type 3:0 Ă, p1 delete 2:4+1, p2 undo, p2 redo, >p2, >p1, >p0, flush',
      'p0 delete 2:0+1, p0 redo, p1 undo, p2 type 1:0 ă, >p2, >p1, >p0, flush',
      'p0 type 1:1 Ą, p0 redo, p1 delete 1:0+1, p1 redo, p2 undo, >p0, >p1, >p2, flush',
    ]
  ),
  undoHistory(
    480,
    [2002, 3003, 1001],
    [
      'p0 enter 0:8, p0 type 1:0 Ā, p1 type 0:6 ā, p1 undo, p2 undo, p2 enter 0:3+1!, >p0, >p2, >p1, flush',
      'p0 redo, p1 redo, p1 redo, p2 redo, >p1, >p2, >p0, flush',
      'p0 delete 2:0+1, p1 enter 3:1+3!, p2 undo, p2 redo, >p2, >p0, >p1, flush',
      'p0 enter 1:0+1!, p0 enter 0:0+3!, p1 undo, p2 redo, >p2, >p0, >p1, flush',
      'p0 type 3:0 Ă, p0 enter 5:0+2!, p1 undo, p2 undo, >p0, >p2, >p1, flush',
      'p0 enter 4:0, p1 undo, p1 undo, p2 redo, p2 redo, >p2, >p0, >p1, flush',
    ]
  ),
  undoHistory(
    597,
    [2002, 1001, 3003],
    [
      'p0 type 1:1 Ā, p0 type 1:4 ā, p1 enter 1:1, p1 type 1:0 Ă, p2 type 1:0 ă, >p2, >p0, >p1, flush',
      'p0 type 0:7 Ą, p1 delete 1:3+1, p1 redo, p2 undo, p2 type 2:0 ą, >p1, >p2, >p0, flush',
      'p0 redo, p1 undo, p2 undo, p2 delete 1:0+1, >p0, >p1, >p2, flush',
      'p0 undo, p0 enter 2:1, p1 undo, p1 type 2:0 Ć, p2 undo, p2 type 2:4 ć, >p0, >p2, >p1, flush',
      'p0 redo, p1 type 2:1 Ĉ, p1 undo, p2 redo, >p1, >p0, >p2, flush',
      'p0 redo, p0 type 3:4 ĉ, p1 type 2:1 Ċ, p2 undo, p2 delete 3:1+1, >p1, >p0, >p2, flush',
    ]
  ),
  undoHistory(
    670,
    [2002, 1001, 3003],
    [
      'p0 delete 1:3+1, p0 enter 0:8, p1 undo, p2 enter 1:1+3!, p2 undo, >p1, >p0, >p2, flush',
      'p0 undo, p1 type 0:1 Ā, p2 type 1:0 ā, p2 type 2:3 Ă, >p1, >p0, >p2, flush',
      'p0 enter 2:3+1!, p0 redo, p1 redo, p2 enter 2:2+2!, >p2, >p1, >p0, flush',
      'p0 type 2:0 ă, p1 undo, p2 redo, >p0, >p2, >p1, flush',
      'p0 type 0:3 Ą, p0 undo, p1 redo, p2 redo, p2 delete 0:2+2, >p0, >p2, >p1, flush',
      'p0 type 3:0 ą, p1 undo, p2 type 3:0 Ć, p2 undo, >p1, >p0, >p2, flush',
    ]
  ),
  undoHistory(
    814,
    [3003, 2002, 1001],
    [
      'p0 type 0:8 Ā, p1 delete 1:3+1, p2 undo, p2 delete 1:3+1, >p1, >p2, >p0, flush',
      'p0 type 0:6 ā, p0 redo, p1 type 0:0 Ă, p1 redo, p2 redo, p2 undo, >p0, >p1, >p2, flush',
      'p0 redo, p1 redo, p2 enter 1:2+1!, >p2, >p0, >p1, flush',
      'p0 undo, p0 undo, p1 redo, p2 enter 1:0+2!, >p1, >p0, >p2, flush',
      'p0 undo, p0 type 0:0 ă, p1 enter 1:0, p1 type 3:0 Ą, p2 undo, p2 type 0:5 ą, >p1, >p2, >p0, flush',
      'p0 delete 1:1+1, p0 type 4:1 Ć, p1 undo, p2 enter 1:0+2!, >p2, >p1, >p0, flush',
    ]
  ),
  undoHistory(
    947,
    [1001, 2002, 3003],
    [
      'p0 type 1:3 Ā, p1 type 1:2 ā, p2 redo, p2 redo, >p2, >p1, >p0, flush',
      'p0 undo, p1 enter 1:5+1!, p2 enter 1:3+3!, >p0, >p1, >p2, flush',
      'p0 redo, p1 type 0:1 Ă, p1 enter 0:0+3!, p2 undo, p2 type 2:0 ă, >p0, >p2, >p1, flush',
      'p0 enter 2:3, p0 redo, p1 type 3:1 Ą, p1 redo, p2 undo, >p1, >p2, >p0, flush',
      'p0 undo, p1 redo, p2 type 1:6 ą, >p1, >p0, >p2, flush',
      'p0 undo, p1 delete 3:0+1, p2 enter 3:0+1!, p2 type 0:0 Ć, >p1, >p2, >p0, flush',
    ]
  ),
  undoHistory(
    1024,
    [2002, 3003, 1001],
    [
      'p0 redo, p1 type 0:4 Ā, p1 enter 1:2+1!, p2 redo, >p2, >p0, >p1, flush',
      'p0 type 0:2 ā, p1 type 0:5 Ă, p2 redo, >p0, >p1, >p2, flush',
      'p0 redo, p0 type 1:0 ă, p1 undo, p1 redo, p2 enter 0:5+2!, >p0, >p2, >p1, flush',
      'p0 type 2:1 Ą, p1 type 0:1 ą, p1 type 0:6 Ć, p2 type 3:0 ć, p2 type 3:1 Ĉ, >p1, >p0, >p2, flush',
      'p0 type 0:7 ĉ, p1 type 1:0 Ċ, p2 delete 1:3+1, >p1, >p0, >p2, flush',
      'p0 type 3:2 ċ, p0 enter 0:5, p1 redo, p1 undo, p2 undo, >p2, >p0, >p1, flush',
    ]
  ),
  undoHistory(
    1129,
    [2002, 3003, 1001],
    [
      'p0 undo, p0 delete 0:7+1, p1 redo, p1 enter 0:8, p2 enter 1:2+2!, >p0, >p1, >p2, flush',
      'p0 type 2:2 Ā, p1 type 3:0 ā, p1 type 3:0 Ă, p2 redo, >p1, >p2, >p0, flush',
      'p0 undo, p0 redo, p1 undo, p2 type 1:0 ă, p2 redo, >p2, >p1, >p0, flush',
      'p0 redo, p0 delete 2:2+1, p1 type 0:2 Ą, p1 delete 0:4+1, p2 undo, >p2, >p0, >p1, flush',
      'p0 type 1:0 ą, p0 delete 3:0+1, p1 type 1:0 Ć, p1 delete 0:2+1, p2 undo, p2 type 3:1 ć, >p0, >p1, >p2, flush',
      'p0 undo, p1 undo, p2 type 3:1 Ĉ, p2 type 2:4 ĉ, >p0, >p2, >p1, flush',
    ]
  ),
  undoHistory(
    1196,
    [2002, 1001, 3003],
    [
      'p0 type 1:2 Ā, p0 type 0:4 ā, p1 type 1:3 Ă, p1 enter 0:6, p2 type 1:0 ă, >p0, >p1, >p2, flush',
      'p0 type 2:1 Ą, p1 redo, p2 type 0:5 ą, >p1, >p0, >p2, flush',
      'p0 type 2:6 Ć, p1 undo, p2 type 0:6 ć, p2 type 1:2 Ĉ, >p2, >p1, >p0, flush',
      'p0 type 1:9 ĉ, p0 enter 1:6+1!, p1 delete 0:2+1, p2 enter 1:7+2!, p2 undo, >p2, >p1, >p0, flush',
      'p0 redo, p0 type 2:3 Ċ, p1 undo, p1 undo, p2 enter 1:0, p2 undo, >p1, >p2, >p0, flush',
      'p0 enter 0:5+3!, p1 redo, p1 type 2:4 ċ, p2 redo, >p1, >p0, >p2, flush',
    ]
  ),
  undoHistory(
    1299,
    [1001, 3003, 2002],
    [
      'p0 redo, p1 delete 1:3+1, p2 undo, p2 redo, >p0, >p2, >p1, flush',
      'p0 type 0:0 Ā, p1 enter 0:5+3!, p1 delete 2:0+2, p2 type 0:4 ā, p2 type 0:2 Ă, >p2, >p0, >p1, flush',
      'p0 redo, p1 redo, p1 type 1:0 ă, p2 undo, >p1, >p0, >p2, flush',
      'p0 type 1:0 Ą, p1 type 1:1 ą, p1 undo, p2 redo, >p1, >p2, >p0, flush',
      'p0 delete 2:0+1, p1 undo, p2 type 1:0 Ć, p2 type 2:0 ć, >p0, >p1, >p2, flush',
      'p0 delete 1:1+1, p0 delete 1:0+1, p1 undo, p2 type 1:2 Ĉ, >p2, >p1, >p0, flush',
    ]
  ),
];

/** Typing, Enter, delete, bold, italic, undo, and redo over 6 rounds: the `FORMAT` mode. */
export const LONG_FORMAT_HISTORIES: readonly FuzzRegression<TextAction>[] = [
  longFormatHistory(
    7,
    [2002, 3003, 1001],
    [
      'p0 type 1:2 Ā, p0 type 1:4 ā, p1 enter 1:0+1!, p1 type 1:0 Ă, p2 enter 0:1, p2 redo, >p0, >p1, >p2, flush',
      'p0 undo, p0 delete 0:0+1, p1 delete 0:0+1, p2 enter 3:4+1!, >p2, >p1, >p0, flush',
      'p0 enter 1:1+2!, p0 bold 0:0+0, p1 bold 3:3+0, p1 undo, p2 undo, p2 italic 0:0+0, >p0, >p1, >p2, flush',
      'p0 italic 2:4+0, p0 type 3:1 ă, p1 italic 0:0+0, p2 redo, >p1, >p0, >p2, flush',
      'p0 type 4:3 Ą, p0 type 4:3 ą, p1 italic 1:1+0, p2 redo, p2 undo, >p0, >p2, >p1, flush',
      'p0 delete 2:0+1, p1 type 0:0 Ć, p2 type 4:6 ć, >p1, >p2, >p0, flush',
    ]
  ),
  longFormatHistory(
    112,
    [1001, 2002, 3003],
    [
      'p0 delete 0:1+2, p0 type 1:3 Ā, p1 delete 0:1+1, p2 type 1:0 ā, >p1, >p0, >p2, flush',
      'p0 undo, p1 undo, p1 italic 1:4+0, p2 undo, >p2, >p0, >p1, flush',
      'p0 bold 1:1+1, p0 bold 1:1+1, p1 bold 1:0+2, p2 enter 0:1, >p0, >p2, >p1, flush',
      'p0 delete 0:0+1, p0 italic 1:4+2, p1 italic 2:2+2, p2 type 2:0 Ă, >p1, >p0, >p2, flush',
      'p0 bold 2:5+0, p1 italic 1:0+0, p2 redo, >p2, >p0, >p1, flush',
      'p0 redo, p0 redo, p1 undo, p2 redo, p2 bold 1:6+0, >p1, >p0, >p2, flush',
    ]
  ),
  longFormatHistory(
    209,
    [2002, 1001, 3003],
    [
      'p0 type 0:3 Ā, p0 redo, p1 enter 1:1, p1 redo, p2 redo, >p0, >p1, >p2, flush',
      'p0 type 2:3 ā, p1 italic 0:7+2, p2 type 2:3 Ă, p2 delete 1:0+1, >p1, >p0, >p2, flush',
      'p0 bold 1:0+0, p0 redo, p1 redo, p1 enter 2:2+1!, p2 delete 2:4+1, p2 redo, >p0, >p2, >p1, flush',
      'p0 italic 2:2+0, p0 enter 1:0, p1 redo, p2 redo, >p1, >p0, >p2, flush',
      'p0 enter 0:8+1!, p0 enter 0:8, p1 delete 3:0+2, p1 enter 2:0, p2 undo, >p2, >p0, >p1, flush',
      'p0 enter 6:0, p1 redo, p2 type 4:0 ă, p2 type 7:2 Ą, >p1, >p2, >p0, flush',
    ]
  ),
  longFormatHistory(
    228,
    [1001, 2002, 3003],
    [
      'p0 enter 1:4, p1 enter 1:4, p2 italic 0:6+2, >p1, >p0, >p2, flush',
      'p0 type 3:0 Ā, p1 bold 0:5+0, p1 enter 0:5+1!, p2 italic 1:0+3, p2 bold 2:0+0, >p1, >p2, >p0, flush',
      'p0 enter 2:3, p0 enter 4:0, p1 undo, p1 bold 2:0+0, p2 undo, p2 type 0:2 ā, >p1, >p0, >p2, flush',
      'p0 delete 5:0+1, p1 bold 5:1+0, p2 italic 0:3+3, p2 delete 2:0+1, >p1, >p0, >p2, flush',
      'p0 enter 2:0, p1 type 4:0 Ă, p2 italic 5:0+0, p2 type 4:0 ă, >p1, >p2, >p0, flush',
      'p0 type 4:0 Ą, p1 undo, p1 redo, p2 redo, >p0, >p1, >p2, flush',
    ]
  ),
  longFormatHistory(
    327,
    [2002, 1001, 3003],
    [
      'p0 type 1:4 Ā, p1 type 0:3 ā, p2 redo, >p0, >p1, >p2, flush',
      'p0 delete 0:6+1, p0 type 0:2 Ă, p1 bold 0:8+1, p1 delete 0:3+3, p2 type 1:1 ă, p2 bold 1:4+1, >p0, >p1, >p2, flush',
      'p0 type 1:5 Ą, p0 type 1:1 ą, p1 redo, p2 redo, p2 type 0:3 Ć, >p1, >p2, >p0, flush',
      'p0 italic 1:7+0, p1 bold 1:3+3, p2 undo, p2 type 1:5 ć, >p1, >p0, >p2, flush',
      'p0 enter 0:2, p1 type 0:3 Ĉ, p1 enter 0:6+1!, p2 redo, p2 enter 0:0+2!, >p0, >p1, >p2, flush',
      'p0 redo, p1 delete 2:2+1, p2 redo, p2 redo, >p0, >p1, >p2, flush',
    ]
  ),
  longFormatHistory(
    363,
    [2002, 1001, 3003],
    [
      'p0 redo, p1 italic 0:7+1, p1 delete 1:1+1, p2 italic 0:5+2, p2 delete 1:1+3, >p1, >p0, >p2, flush',
      'p0 bold 0:1+2, p0 type 1:1 Ā, p1 enter 1:0, p1 undo, p2 enter 1:1, >p1, >p0, >p2, flush',
      'p0 type 1:0 ā, p0 italic 0:7+0, p1 bold 2:0+0, p1 undo, p2 italic 1:0+0, p2 italic 2:0+0, >p1, >p2, >p0, flush',
      'p0 enter 2:0, p0 type 3:0 Ă, p1 type 2:0 ă, p2 italic 1:0+1, p2 redo, >p2, >p1, >p0, flush',
      'p0 delete 2:0+1, p0 redo, p1 bold 2:0+0, p2 type 1:2 Ą, >p2, >p1, >p0, flush',
      'p0 type 1:3 ą, p0 redo, p1 delete 1:2+1, p1 redo, p2 bold 1:4+0, >p0, >p1, >p2, flush',
    ]
  ),
  longFormatHistory(
    480,
    [2002, 3003, 1001],
    [
      'p0 enter 0:8, p0 type 1:0 Ā, p1 type 0:6 ā, p1 bold 0:8+1, p2 bold 1:3+1, p2 enter 0:3+1!, >p0, >p2, >p1, flush',
      'p0 redo, p1 undo, p1 redo, p2 redo, >p1, >p2, >p0, flush',
      'p0 delete 2:0+1, p1 enter 3:1+3!, p2 italic 0:3+0, p2 redo, >p2, >p0, >p1, flush',
      'p0 enter 1:0+1!, p0 enter 0:0+3!, p1 italic 2:0+0, p2 redo, >p2, >p0, >p1, flush',
      'p0 type 3:0 Ă, p0 enter 5:0+1!, p1 italic 1:1+0, p2 bold 3:3+0, >p0, >p2, >p1, flush',
      'p0 enter 6:0, p1 bold 3:0+2, p1 bold 1:1+0, p2 undo, p2 undo, >p2, >p0, >p1, flush',
    ]
  ),
  longFormatHistory(
    670,
    [2002, 1001, 3003],
    [
      'p0 delete 1:3+1, p0 enter 0:8, p1 italic 1:0+0, p2 enter 1:1+3!, p2 bold 0:6+2, >p1, >p0, >p2, flush',
      'p0 italic 3:0+0, p1 type 0:1 Ā, p2 type 1:0 ā, p2 type 2:1 Ă, >p1, >p0, >p2, flush',
      'p0 enter 3:0, p0 redo, p1 undo, p2 enter 3:0, >p2, >p1, >p0, flush',
      'p0 type 3:0 ă, p1 bold 4:0+0, p2 redo, >p0, >p2, >p1, flush',
      'p0 type 0:3 Ą, p0 bold 5:0+0, p1 undo, p2 undo, p2 delete 0:2+2, >p0, >p2, >p1, flush',
      'p0 type 4:0 ą, p1 bold 4:0+0, p2 type 4:0 Ć, p2 italic 3:0+1, >p1, >p0, >p2, flush',
    ]
  ),
  longFormatHistory(
    814,
    [3003, 2002, 1001],
    [
      'p0 type 0:8 Ā, p1 delete 1:3+1, p2 italic 0:7+1, p2 delete 1:3+1, >p1, >p2, >p0, flush',
      'p0 type 0:6 ā, p0 redo, p1 type 0:0 Ă, p1 undo, p2 redo, p2 italic 1:2+1, >p0, >p1, >p2, flush',
      'p0 redo, p1 undo, p2 enter 1:2+1!, >p2, >p0, >p1, flush',
      'p0 bold 0:2+1, p0 bold 0:6+2, p1 redo, p2 enter 1:0+3!, >p1, >p0, >p2, flush',
      'p0 italic 2:0+0, p0 type 0:0 ă, p1 enter 1:0, p1 type 3:0 Ą, p2 italic 1:0+0, p2 type 0:6 ą, >p1, >p2, >p0, flush',
      'p0 type 1:0 Ć, p0 type 4:0 ć, p1 italic 2:0+0, p2 enter 1:0, >p2, >p1, >p0, flush',
    ]
  ),
  longFormatHistory(
    947,
    [1001, 2002, 3003],
    [
      'p0 type 1:3 Ā, p1 type 1:2 ā, p2 undo, p2 undo, >p2, >p1, >p0, flush',
      'p0 italic 1:5+1, p1 enter 1:5+1!, p2 enter 1:3+3!, >p0, >p1, >p2, flush',
      'p0 redo, p1 type 0:1 Ă, p1 enter 0:0+3!, p2 italic 0:5+2, p2 type 3:0 ă, >p0, >p2, >p1, flush',
      'p0 enter 2:1, p0 redo, p1 type 4:1 Ą, p1 undo, p2 bold 1:4+0, >p1, >p2, >p0, flush',
      'p0 bold 3:0+1, p1 undo, p2 type 1:4 ą, >p1, >p0, >p2, flush',
      'p0 bold 0:7+2, p1 type 4:0 Ć, p2 enter 5:0+1!, p2 delete 0:0+1, >p1, >p2, >p0, flush',
    ]
  ),
  longFormatHistory(
    1024,
    [2002, 3003, 1001],
    [
      'p0 undo, p1 type 0:4 Ā, p1 enter 1:2+1!, p2 redo, >p2, >p0, >p1, flush',
      'p0 type 0:2 ā, p1 type 0:5 Ă, p2 redo, >p0, >p1, >p2, flush',
      'p0 redo, p0 type 1:0 ă, p1 italic 2:0+1, p1 redo, p2 enter 0:5+2!, >p0, >p2, >p1, flush',
      'p0 type 2:1 Ą, p1 type 0:1 ą, p1 type 0:5 Ć, p2 type 3:0 ć, p2 type 3:1 Ĉ, >p1, >p0, >p2, flush',
      'p0 type 0:6 ĉ, p1 type 1:0 Ċ, p2 delete 1:3+1, >p1, >p0, >p2, flush',
      'p0 type 3:2 ċ, p0 enter 0:5, p1 redo, p1 italic 3:0+2, p2 italic 0:0+3, >p2, >p0, >p1, flush',
    ]
  ),
  longFormatHistory(
    1129,
    [2002, 3003, 1001],
    [
      'p0 bold 0:2+2, p0 delete 0:7+1, p1 redo, p1 enter 0:8, p2 enter 1:2+2!, >p0, >p1, >p2, flush',
      'p0 type 2:2 Ā, p1 type 3:0 ā, p1 type 3:0 Ă, p2 undo, >p1, >p2, >p0, flush',
      'p0 bold 3:0+2, p0 redo, p1 italic 3:1+1, p2 type 1:0 ă, p2 redo, >p2, >p1, >p0, flush',
      'p0 redo, p0 delete 2:3+2, p1 type 0:2 Ą, p1 delete 0:4+1, p2 bold 2:1+3, >p2, >p0, >p1, flush',
      'p0 type 1:0 ą, p0 delete 3:1+1, p1 delete 1:0+1, p1 delete 0:2+1, p2 italic 1:1+0, p2 type 3:2 Ć, >p0, >p1, >p2, flush',
      'p0 italic 2:0+0, p1 italic 3:2+0, p2 type 3:2 ć, p2 type 2:3 Ĉ, >p0, >p2, >p1, flush',
    ]
  ),
  longFormatHistory(
    1205,
    [2002, 1001, 3003],
    [
      'p0 italic 0:1+2, p0 italic 0:2+2, p1 redo, p1 enter 0:3, p2 enter 1:3, >p0, >p1, >p2, flush',
      'p0 bold 2:0+0, p0 type 1:0 Ā, p1 redo, p2 undo, p2 enter 0:2, >p1, >p2, >p0, flush',
      'p0 delete 0:0+2, p1 type 1:0 ā, p1 italic 1:0+1, p2 italic 3:0+2, >p1, >p2, >p0, flush',
      'p0 undo, p1 delete 2:4+2, p1 italic 3:1+1, p2 bold 3:4+0, >p0, >p1, >p2, flush',
      'p0 italic 3:1+2, p0 italic 3:2+2, p1 redo, p1 type 1:1 Ă, p2 bold 2:3+1, >p0, >p1, >p2, flush',
      'p0 bold 2:2+0, p0 delete 3:3+1, p1 italic 1:0+3, p2 type 2:3 ă, p2 bold 1:0+2, >p1, >p0, >p2, flush',
    ]
  ),
  longFormatHistory(
    1299,
    [1001, 3003, 2002],
    [
      'p0 undo, p1 delete 1:3+1, p2 bold 0:3+1, p2 undo, >p0, >p2, >p1, flush',
      'p0 type 0:0 Ā, p1 enter 0:5+3!, p1 delete 2:0+2, p2 type 0:4 ā, p2 type 0:2 Ă, >p2, >p0, >p1, flush',
      'p0 redo, p1 redo, p1 type 1:0 ă, p2 bold 0:1+2, >p1, >p0, >p2, flush',
      'p0 type 1:0 Ą, p1 type 1:1 ą, p1 italic 0:4+1, p2 undo, >p1, >p2, >p0, flush',
      'p0 delete 2:0+1, p1 bold 1:3+0, p2 type 1:1 Ć, p2 type 2:0 ć, >p0, >p1, >p2, flush',
      'p0 delete 1:3+1, p0 delete 1:2+1, p1 italic 2:1+0, p2 type 1:4 Ĉ, >p2, >p1, >p0, flush',
    ]
  ),
  longFormatHistory(
    1919,
    [1001, 3003, 2002],
    [
      'p0 undo, p1 type 1:3 Ā, p2 bold 0:1+3, p2 bold 0:3+1, >p2, >p1, >p0, flush',
      'p0 undo, p1 type 0:6 ā, p2 type 1:3 Ă, p2 enter 0:1, >p1, >p2, >p0, flush',
      'p0 type 1:6 ă, p0 italic 2:5+1, p1 enter 1:1+1!, p2 enter 0:1, p2 italic 0:1+0, >p0, >p2, >p1, flush',
      'p0 delete 3:2+1, p0 type 1:0 Ą, p1 type 3:4 ą, p1 undo, p2 italic 1:0+0, >p0, >p2, >p1, flush',
      'p0 type 2:1 Ć, p0 delete 2:0+1, p1 enter 2:0+1!, p2 redo, >p2, >p0, >p1, flush',
      'p0 bold 0:0+0, p0 italic 2:1+0, p1 redo, p2 delete 0:0+1, p2 italic 3:0+0, >p0, >p2, >p1, flush',
    ]
  ),
];

/**
 * Of those seeds, the histories that fail under the stricter checks: an undo or a redo brings
 * back a character that a peer removed (#3).
 */
export const UNDO_HISTORY_FAILURES: readonly FuzzKnownFailure<TextAction>[] = [
  failing(
    undoHistory(
      733,
      [2002, 1001, 3003],
      [
        'p0 enter 1:2, p1 redo, p1 type 1:2 Ā, p2 type 1:1 ā, >p0, >p2, >p1, flush',
        'p0 undo, p1 undo, p1 delete 2:1+1, p2 redo, p2 undo, >p1, >p0, >p2, flush',
        'p0 redo, p1 delete 1:1+3, p2 type 1:4 Ă, >p1, >p0, >p2, flush',
        'p0 undo, p0 delete 0:0+3, p1 undo, p2 type 2:1 ă, p2 type 1:2 Ą, >p1, >p0, >p2, flush',
        'p0 undo, p1 redo, p2 type 0:2 ą, p2 delete 0:0+2, >p0, >p1, >p2, flush',
        'p0 enter 0:0+1!, p0 delete 1:5+1, p1 enter 2:0, p2 delete 1:4+1, >p0, >p1, >p2, flush',
      ]
    ),
    '#3',
    'removed characters show again: "L" in ["ABCDEFGH","IJKL"]'
  ),
  failing(
    undoHistory(
      1205,
      [2002, 1001, 3003],
      [
        'p0 undo, p0 undo, p1 redo, p1 enter 0:3, p2 enter 1:3, >p0, >p1, >p2, flush',
        'p0 undo, p0 type 1:0 Ā, p1 redo, p2 redo, p2 enter 0:2, >p1, >p2, >p0, flush',
        'p0 delete 1:0+1, p1 type 2:0 ā, p1 undo, p2 undo, >p1, >p2, >p0, flush',
        'p0 redo, p1 delete 2:2+1, p1 undo, p2 undo, >p0, >p1, >p2, flush',
        'p0 undo, p0 undo, p1 redo, p1 type 1:4 Ă, p2 undo, >p0, >p1, >p2, flush',
        'p0 undo, p0 delete 3:2+1, p1 undo, p2 type 2:5 ă, p2 undo, >p1, >p0, >p2, flush',
      ]
    ),
    '#3',
    'removed characters show again: "C" in ["ABC","ĀDEFGH","IJK","L"]'
  ),
  failing(
    undoHistory(
      1919,
      [1001, 3003, 2002],
      [
        'p0 redo, p1 type 1:3 Ā, p2 undo, p2 undo, >p2, >p1, >p0, flush',
        'p0 redo, p1 type 0:6 ā, p2 type 1:3 Ă, p2 enter 0:1, >p1, >p2, >p0, flush',
        'p0 type 1:6 ă, p0 undo, p1 enter 1:1+1!, p2 enter 0:1, p2 undo, >p0, >p2, >p1, flush',
        'p0 delete 2:2+1, p0 type 0:1 Ą, p1 type 2:3 ą, p1 redo, p2 undo, >p0, >p2, >p1, flush',
        'p0 type 1:1 Ć, p0 delete 1:0+1, p1 enter 1:0+1!, p2 redo, >p2, >p0, >p1, flush',
        'p0 undo, p0 undo, p1 redo, p2 delete 0:1+1, p2 undo, >p0, >p2, >p1, flush',
      ]
    ),
    '#3',
    'p2 undo added characters it may not add: "C"'
  ),
];

/** The same, with formats. */
export const LONG_FORMAT_HISTORY_FAILURES: readonly FuzzKnownFailure<TextAction>[] = [
  failing(
    longFormatHistory(
      304,
      [2002, 3003, 1001],
      [
        'p0 italic 1:3+0, p0 bold 0:4+1, p1 enter 0:1+2!, p2 bold 1:1+2, >p2, >p0, >p1, flush',
        'p0 delete 2:1+2, p0 type 1:1 Ā, p1 italic 2:4+0, p2 type 1:2 ā, p2 redo, >p0, >p2, >p1, flush',
        'p0 delete 1:1+2, p0 type 0:2 Ă, p1 bold 0:1+0, p1 italic 1:4+2, p2 redo, >p2, >p0, >p1, flush',
        'p0 undo, p1 bold 0:1+1, p1 type 1:4 ă, p2 type 0:3 Ą, >p2, >p1, >p0, flush',
        'p0 italic 2:0+2, p0 redo, p1 type 2:0 ą, p1 type 1:0 Ć, p2 undo, >p1, >p0, >p2, flush',
        'p0 redo, p1 type 2:3 ć, p1 type 2:0 Ĉ, p2 redo, >p2, >p1, >p0, flush',
      ]
    ),
    '#3',
    'removed characters show again: "Ă" in ["AFĂĄ","DāGHă","IL"]'
  ),
  failing(
    longFormatHistory(
      597,
      [2002, 1001, 3003],
      [
        'p0 type 1:1 Ā, p0 type 1:4 ā, p1 enter 1:1, p1 type 1:0 Ă, p2 type 1:0 ă, >p2, >p0, >p1, flush',
        'p0 type 0:7 Ą, p1 delete 1:3+1, p1 redo, p2 italic 0:0+3, p2 type 2:0 ą, >p1, >p2, >p0, flush',
        'p0 redo, p1 italic 1:1+2, p2 italic 1:1+1, p2 delete 1:1+1, >p0, >p1, >p2, flush',
        'p0 italic 0:2+1, p0 enter 2:2, p1 bold 0:0+2, p1 type 2:0 Ć, p2 bold 1:0+1, p2 type 2:5 ć, >p0, >p2, >p1, flush',
        'p0 redo, p1 type 2:2 Ĉ, p1 bold 3:1+1, p2 undo, >p1, >p0, >p2, flush',
        'p0 undo, p0 type 3:3 ĉ, p1 type 2:2 Ċ, p2 italic 3:3+1, p2 delete 3:2+1, >p1, >p0, >p2, flush',
      ]
    ),
    '#3',
    'removed characters show again: "ć" in ["ABCDEFGĄH","ĂI","ĆąĈJ","KāLć"]'
  ),
  failing(
    longFormatHistory(
      733,
      [2002, 1001, 3003],
      [
        'p0 enter 1:2, p1 undo, p1 type 1:2 Ā, p2 type 1:1 ā, >p0, >p2, >p1, flush',
        'p0 bold 2:2+0, p1 italic 1:1+1, p1 delete 2:1+1, p2 undo, p2 italic 0:2+1, >p1, >p0, >p2, flush',
        'p0 undo, p1 delete 2:0+1, p2 type 2:1 Ă, >p1, >p0, >p2, flush',
        'p0 bold 2:1+0, p0 delete 1:0+3, p1 bold 0:2+1, p2 type 2:0 ă, p2 type 1:4 Ą, >p1, >p0, >p2, flush',
        'p0 bold 0:1+2, p1 redo, p2 type 0:4 ą, p2 delete 0:1+2, >p0, >p1, >p2, flush',
        'p0 enter 0:0+1!, p0 delete 1:5+1, p1 enter 2:1, p2 delete 1:3+1, >p0, >p1, >p2, flush',
      ]
    ),
    '#3',
    'removed characters show again: "ā" in ["ABCDEFGH","IāJĀ","K"]'
  ),
  failing(
    longFormatHistory(
      1196,
      [2002, 1001, 3003],
      [
        'p0 type 1:2 Ā, p0 type 0:4 ā, p1 type 1:3 Ă, p1 enter 0:6, p2 type 1:0 ă, >p0, >p1, >p2, flush',
        'p0 type 2:1 Ą, p1 redo, p2 type 0:5 ą, >p1, >p0, >p2, flush',
        'p0 type 2:6 Ć, p1 italic 0:2+3, p2 type 0:6 ć, p2 type 1:2 Ĉ, >p2, >p1, >p0, flush',
        'p0 type 1:3 ĉ, p0 enter 2:6+1!, p1 delete 0:1+1, p2 enter 2:7+2!, p2 italic 0:4+2, >p2, >p1, >p0, flush',
        'p0 undo, p0 type 3:0 Ċ, p1 bold 2:6+0, p1 italic 4:0+0, p2 enter 1:0, p2 italic 2:2+2, >p1, >p2, >p0, flush',
        'p0 enter 0:3+3!, p1 redo, p1 type 4:1 ċ, p2 redo, >p1, >p0, >p2, flush',
      ]
    ),
    '#3',
    'p0 undo added characters it may not add: "ĂL"'
  ),
];
