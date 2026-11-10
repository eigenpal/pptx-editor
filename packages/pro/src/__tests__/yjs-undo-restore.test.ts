/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The characters an undo or a redo puts back (`undo-restore.ts`, the restore marks of
// `text-identity.ts`) and the nodes an undo takes out (`undo-unlisted.ts`), with two or three
// editors on linked Y.Docs. Each case checks the text, not only that the replicas converge.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { RESTORED_KEY } from '../collaboration/yjs/text-identity.ts';
import { UNDONE_KEY } from '../collaboration/yjs/undo-unlisted.ts';
import {
  all,
  converged,
  IN,
  oneRun,
  peers,
  replicas,
  root,
  selectIn,
  slide1,
  transact,
} from './lineage-support.ts';

const ORDERS = [false, true] as const;
const order = (swap: boolean) => `client order ${swap ? 'B, A' : 'A, B'}`;
const THREE: readonly (readonly number[])[] = [
  [1_001, 2_002, 3_003],
  [3_003, 2_002, 1_001],
];
const FOUR: readonly (readonly number[])[] = [
  ...THREE,
  [2_002, 1_001, 3_003],
  [1_001, 3_003, 2_002],
];
const ok = (result: { ok: boolean }) => expect(result.ok).toBe(true);

const textBox = (body: string) =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text: body }),
      },
    ],
  });

describe('a peer edits text that an undo put back', () => {
  for (const swap of ORDERS)
    for (const [from, to, want] of [
      [1, 3, ['A', 'DEFG']],
      [1, 1, ['A', 'BCDEFG']],
      [3, 3, ['ABC', 'DEFG']],
      [4, 4, ['ABCD', 'EFG']],
    ] as const)
      test(`an Enter at ${from}..${to} keeps every character (${order(swap)})`, async () => {
        const { a, b } = await replicas(oneRun('ABCDEFG'), swap);
        selectIn(a, 0, 1, 4);
        ok(a.input('deleteBackward'));
        ok(a.exec({ type: 'undo' }));
        selectIn(b, 0, from, to);
        ok(b.input('paragraph'));
        expect(converged(a, b)).toEqual([...want]);
      });

  for (const swap of ORDERS)
    test(`a concurrent delete survives an undo of Enter (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('IJKL'), swap);
      sync.pause();
      selectIn(a, 0, 3, 4);
      ok(a.input('deleteBackward'));
      selectIn(b, 0, 1, 4);
      ok(b.input('paragraph'));
      ok(b.exec({ type: 'undo' }));
      sync.flush();
      expect(converged(a, b)).toEqual(['IJK']);
      // Later splits keep the independently deleted character absent.
      sync.pause();
      selectIn(a, 0, 3);
      ok(a.input('paragraph'));
      selectIn(b, 0, 2, 3);
      ok(b.input('paragraph'));
      sync.flush();
      const text = converged(a, b).join('');
      expect(text).not.toContain('L');
      expect(text.length).toBe(new Set(text).size);
    });
});

describe('a character two undos put back shows once', () => {
  for (const ids of THREE)
    test(`two peers undo their delete of one character at once (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(oneRun('ABCDEFGH'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      selectIn(a, 0, 0, 1);
      ok(a.input('deleteBackward'));
      selectIn(b, 0, 0, 1);
      ok(b.input('deleteBackward'));
      sync.flush();
      expect(converged(...editors)).toEqual(['BCDEFGH']);
      sync.pause();
      ok(a.exec({ type: 'undo' }));
      ok(b.exec({ type: 'undo' }));
      sync.flush();
      expect(converged(...editors)).toEqual(['ABCDEFGH']);
    });

  for (const ids of THREE)
    test(`an undo after a peer's undo put the same character back (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(oneRun('ABCDEFGH'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      selectIn(a, 0, 1, 3);
      ok(a.input('deleteBackward'));
      selectIn(b, 0, 2);
      ok(b.input('deleteBackward'));
      sync.flush();
      ok(b.exec({ type: 'undo' }));
      expect(converged(...editors)).toEqual(['ABDEFGH']);
      ok(a.exec({ type: 'undo' }));
      expect(converged(...editors)).toEqual(['ABCDEFGH']);
    });

  for (const ids of THREE)
    test(`an undo while a peer's Enter moves the same text keeps it once and in order (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(oneRun('ABCDEFGH'), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      selectIn(a, 0, 1, 3);
      ok(a.input('deleteBackward'));
      selectIn(b, 0, 2);
      ok(b.input('deleteBackward'));
      sync.flush();
      ok(b.exec({ type: 'undo' }));
      sync.pause();
      ok(a.exec({ type: 'undo' }));
      selectIn(c, 0, 1);
      ok(c.input('paragraph'));
      sync.flush();
      expect(converged(...editors)).toEqual(['A', 'BCDEFGH']);
    });

  // The undoing replica receives the Enter first: its Follow repair takes what its undo put
  // back with the copy, and a character whose root the copy already holds goes ("Once").
  for (const ids of FOUR)
    test(`an Enter that reaches the undoing peer first keeps the text once (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync, docs } = await peers(oneRun('ABCDEFGH'), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      selectIn(a, 0, 1, 3);
      ok(a.input('deleteBackward'));
      selectIn(b, 0, 2);
      ok(b.input('deleteBackward'));
      sync.flush();
      ok(b.exec({ type: 'undo' }));
      sync.pause();
      ok(a.exec({ type: 'undo' }));
      selectIn(c, 0, 1);
      ok(c.input('paragraph'));
      sync.deliverTo(docs[0]!);
      sync.flush();
      expect(converged(...editors)).toEqual(['A', 'BCDEFGH']);
    });

  // The undo of the Enter does not put back the character that the peer deleted at the same
  // time (`undo-moves.ts`). The peer's undo of its delete then puts it back in the Enter's
  // copy, and it shows once.
  for (const ids of FOUR)
    test(`an undo of an Enter and an undo of a delete of its copy show it once (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(oneRun('ABCDEFGH'), ids);
      const [a, , c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      selectIn(c, 0, 2);
      ok(c.input('paragraph'));
      sync.pause();
      selectIn(a, 1, 0, 1);
      ok(a.input('deleteBackward'));
      ok(c.exec({ type: 'undo' }));
      sync.flush();
      expect(converged(...editors)).toEqual(['ABDEFGH']);
      ok(a.exec({ type: 'undo' }));
      const text = converged(...editors).join('');
      expect([...text].sort().join('')).toBe('ABCDEFGH');
    });
});

describe('a node an undo took out', () => {
  for (const swap of ORDERS)
    test(`a delete of the peer's text in it leaves the replicas alike (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello'), swap);
      selectIn(a, 0, 5);
      ok(a.input('paragraph'));
      selectIn(b, 1, 0);
      ok(b.input('text', 'Z'));
      ok(a.exec({ type: 'undo' }));
      expect(converged(a, b)).toEqual(['Hello', 'Z']);
      selectIn(b, 1, 0, 1);
      ok(b.input('deleteBackward'));
      expect(converged(a, b)).toEqual(['Hello']);
    });

  for (const ids of THREE)
    test(`a run two peers remove at once keeps the text a third peer typed (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(
        textBox(para(run('AB')) + para('<a:endParaRPr lang="en-US"/>')),
        ids
      );
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      selectIn(c, 1, 0);
      ok(c.input('text', 'x'));
      sync.pause();
      selectIn(a, 1, 1);
      ok(a.input('text', 'y'));
      selectIn(b, 1, 1);
      ok(b.input('deleteBackward'));
      ok(c.exec({ type: 'undo' }));
      sync.flush();
      expect(converged(...editors)).toEqual(['AB', 'y']);
    });

  for (const swap of ORDERS)
    test(`it goes before its next sibling when a peer removed the one before it (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(
        textBox(para(run('One')) + para(run('Two')) + para(run('Three'))),
        swap
      );
      selectIn(a, 0, 3);
      ok(a.input('paragraph'));
      selectIn(b, 1, 0);
      ok(b.input('text', 'x'));
      expect(converged(a, b)).toEqual(['One', 'x', 'Two', 'Three']);
      sync.pause();
      ok(a.exec({ type: 'undo' }));
      transact(b, [{ op: 'removeNode', part: slide1(b), node: all(root(b), 'p')[0]!.id }]);
      sync.flush();
      expect(converged(a, b)).toEqual(['x', 'Two', 'Three']);
    });

  for (const swap of ORDERS)
    test(`it still shows after more undone listings than any cap (${order(swap)})`, async () => {
      const { a, b, docs } = await replicas(oneRun('Hello'), swap);
      // A long session: 25,000 listings of other nodes of the same part, never redone.
      docs[0].transact(() => {
        const undone = docs[0].getMap(UNDONE_KEY);
        for (let i = 0; i < 25_000; i += 1)
          undone.set(`gone${i}`, { p: slide1(a), u: 'none', a: '', n: '', c: 7, q: 0 });
      });
      selectIn(a, 0, 5);
      ok(a.input('paragraph'));
      selectIn(b, 1, 0);
      ok(b.input('text', 'Z'));
      ok(a.exec({ type: 'undo' }));
      expect(converged(a, b)).toEqual(['Hello', 'Z']);
      // A write of the parent keeps it.
      selectIn(a, 0, 5);
      ok(a.input('text', '!'));
      expect(converged(a, b)).toEqual(['Hello!', 'Z']);
    });
});

describe('restore marks', () => {
  for (const swap of ORDERS)
    test(`each undo or redo adds its own marks and changes none before (${order(swap)})`, async () => {
      const { a, b, docs } = await replicas(oneRun('Hello world'), swap);
      const marks = docs[0].getMap(RESTORED_KEY);
      selectIn(a, 0, 0, 5);
      ok(a.input('deleteBackward'));
      let before = new Map<string, unknown>();
      for (let i = 0; i < 4; i += 1) {
        ok(a.exec({ type: i % 2 === 0 ? 'undo' : 'redo' }));
        if (i % 2 === 0) {
          expect(marks.size).toBeGreaterThan(before.size);
          for (const [k, v] of before) expect(marks.get(k)).toEqual(v);
          before = new Map(marks.entries());
        }
      }
      expect(converged(a, b)).toEqual([' world']);
    });

  for (const swap of ORDERS)
    test(`a forged mark is ignored or only moves the character it names (${order(swap)})`, async () => {
      const { a, b, docs } = await replicas(oneRun('ABCDEFGH'), swap);
      docs[1].transact(() => {
        const marks = docs[1].getMap(RESTORED_KEY);
        marks.set('__proto__', [1, 1, 1]);
        marks.set('x:1', [1, 1, 1]);
        marks.set(`${docs[1].clientID}:0`, [1, 2, 2 ** 40]);
        marks.set(`${docs[1].clientID}:1`, 'no');
      });
      expect(docs[0].getMap(RESTORED_KEY).size).toBe(4);
      selectIn(a, 0, 1, 3);
      ok(a.input('deleteBackward'));
      ok(a.exec({ type: 'undo' }));
      selectIn(b, 0, 2);
      ok(b.input('paragraph'));
      expect(converged(a, b)).toEqual(['AB', 'CDEFGH']);
      expect(Object.prototype).not.toHaveProperty('1');
    });
});
