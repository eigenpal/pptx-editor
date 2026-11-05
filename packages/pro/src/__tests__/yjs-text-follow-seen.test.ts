/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The "Follow" repair (openspec change `collaboration-text-lineage`, "Character identity") moves
// only text that the split's writer had not seen. A character that the writer saw deleted stays
// deleted after its Enter, also when the delete reaches a replica after the Enter does
// (eigenpal/pptx-editor-miro#144). A character that the writer had not seen follows the Enter; when
// a peer had deleted it, that peer's "Delete" repair reaches the moved character.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { IDENTITY, followsOf, seenBy } from '../collaboration/yjs/text-identity.ts';
import { IN, converged, paragraphs, peers, selectIn } from './lineage-support.ts';

/** Three runs and a paragraph end, then a second paragraph. */
const deck = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 2 * IN,
          text:
            para(run('ABCD') + run('EF', 'b="1"') + run('GH') + '<a:endParaRPr lang="en-US"/>') +
            para(run('IJKL')),
        }),
      },
    ],
  });

/**
 * One queue per link, as a transport keeps them, so a test chooses the order in which a
 * replica receives the updates of two peers.
 */
function links(docs: readonly Y.Doc[]) {
  const queues = new Map<string, Uint8Array[]>();
  const queue = (from: number, to: number): Uint8Array[] => {
    const key = `${from}:${to}`;
    let q = queues.get(key);
    if (!q) queues.set(key, (q = []));
    return q;
  };
  for (const [from, doc] of docs.entries())
    doc.on('update', (update: Uint8Array, origin: unknown) => {
      if (origin === queues) return;
      for (const to of docs.keys()) if (to !== from) queue(from, to).push(update);
    });
  const deliver = (from: number, to: number): void => {
    for (const update of queue(from, to).splice(0)) Y.applyUpdate(docs[to]!, update, queues);
  };
  return {
    deliver,
    /** Deliver every waiting update, link by link: sender first, then receiver. */
    flush(): void {
      for (let more = true; more;) {
        more = false;
        for (const from of docs.keys())
          for (const to of docs.keys())
            if (from !== to && queue(from, to).length > 0) {
              more = true;
              deliver(from, to);
            }
      }
    },
  };
}

/** Editors on linked Y.Docs whose updates wait in `links` until a test delivers them. */
async function held(ids: readonly number[]) {
  const { editors, sync, docs } = await peers(deck(), ids);
  // The mesh holds its copies for ever; `links` delivers instead.
  sync.pause();
  return { editors, relay: links(docs) };
}

const THREE_ORDERS = [
  [1_001, 2_002, 3_003],
  [3_003, 2_002, 1_001],
  [2_002, 3_003, 1_001],
];
const TWO_ORDERS = [
  [1_001, 2_002],
  [2_002, 1_001],
];
const order = (ids: readonly number[]) => `client ids ${ids.join(', ')}`;

describe('a deleted character stays deleted after an Enter by a peer that saw the delete', () => {
  for (const ids of THREE_ORDERS)
    test(`a third peer deletes it (${order(ids)})`, async () => {
      const { editors, relay } = await held(ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 8);
      expect(a.input('text', 'x').ok).toBe(true);
      relay.flush();
      selectIn(c, 0, 3, 4);
      expect(c.input('deleteBackward').ok).toBe(true);
      relay.deliver(2, 1);
      expect(paragraphs(b)).toEqual(['ABCEFGHx', 'IJKL']);
      selectIn(b, 0, 1);
      expect(b.input('paragraph').ok).toBe(true);
      expect(paragraphs(b)).toEqual(['A', 'BCEFGHx', 'IJKL']);
      // A receives the Enter before the delete.
      relay.flush();
      expect(converged(a, b, c)).toEqual(['A', 'BCEFGHx', 'IJKL']);
    });

  for (const ids of TWO_ORDERS)
    test(`the peer that presses Enter deletes it (${order(ids)})`, async () => {
      const { editors, relay } = await held(ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 8);
      expect(a.input('text', 'x').ok).toBe(true);
      relay.flush();
      selectIn(b, 0, 3, 4);
      expect(b.input('deleteBackward').ok).toBe(true);
      selectIn(b, 0, 1);
      expect(b.input('paragraph').ok).toBe(true);
      relay.flush();
      expect(converged(a, b)).toEqual(['A', 'BCEFGHx', 'IJKL']);
    });

  for (const ids of THREE_ORDERS)
    test(`text typed after it at the same time still follows the Enter (${order(ids)})`, async () => {
      const { editors, relay } = await held(ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 3);
      expect(a.input('text', 'Z').ok).toBe(true);
      relay.flush();
      // C deletes "Z"; B sees the delete, A types "W" after "Z" and sees neither.
      selectIn(c, 0, 3, 4);
      expect(c.input('deleteBackward').ok).toBe(true);
      relay.deliver(2, 1);
      selectIn(a, 0, 4);
      expect(a.input('text', 'W').ok).toBe(true);
      selectIn(b, 0, 1);
      expect(b.input('paragraph').ok).toBe(true);
      expect(paragraphs(b)).toEqual(['A', 'BCDEFGH', 'IJKL']);
      relay.flush();
      expect(converged(a, b, c)).toEqual(['A', 'BCWDEFGH', 'IJKL']);
    });
});

describe('a deleted character stays deleted after a peer moved it into an Enter', () => {
  // The peer that moves the character has not seen the delete yet; the peer that presses
  // Enter has not seen the character. The deleter receives the copy and the move in either order.
  const DELIVERIES = [
    { name: 'the copy first', first: 1, then: 0 },
    { name: 'the move first', first: 0, then: 1 },
  ];
  /** A types "Z" that only C receives and deletes; A moves it along with B's Enter. */
  async function moveBeforeDelete(ids: readonly number[]) {
    const { editors, relay } = await held(ids);
    const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
    selectIn(a, 0, 3);
    expect(a.input('text', 'Z').ok).toBe(true);
    relay.deliver(0, 2);
    selectIn(c, 0, 3, 4);
    expect(c.input('deleteBackward').ok).toBe(true);
    // B presses Enter after "A". A receives it before C's delete and moves "Z" along.
    selectIn(b, 0, 1);
    expect(b.input('paragraph').ok).toBe(true);
    relay.deliver(1, 0);
    expect(paragraphs(a)).toEqual(['A', 'BCZDEFGH', 'IJKL']);
    return { a, b, c, relay };
  }

  for (const ids of THREE_ORDERS)
    for (const delivery of DELIVERIES)
      test(`the deleter gets ${delivery.name} (${order(ids)})`, async () => {
        const { a, b, c, relay } = await moveBeforeDelete(ids);
        relay.deliver(delivery.first, 2);
        relay.deliver(delivery.then, 2);
        relay.flush();
        expect(converged(a, b, c)).toEqual(['A', 'BCDEFGH', 'IJKL']);
      });

  for (const ids of THREE_ORDERS)
    test(`the deleter's undo shows the character once, and redo deletes it (${order(ids)})`, async () => {
      const { a, b, c, relay } = await moveBeforeDelete(ids);
      relay.flush();
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      relay.flush();
      expect(converged(a, b, c).join('').match(/Z/g)).toHaveLength(1);
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      relay.flush();
      expect(converged(a, b, c)).toEqual(['A', 'BCDEFGH', 'IJKL']);
    });
});

describe('the marks of a copy', () => {
  test('a record with no mark or a wrong one has no seen test', () => {
    const doc = new Y.Doc();
    const rec = doc.getMap<unknown>('r');
    expect(seenBy(rec)).toBeNull();
    for (const wrong of [[1], [1, -1], [1, 2, 1, 3], ['1', 2], 'x'] as unknown[]) {
      rec.set(IDENTITY.seen, wrong);
      expect(seenBy(rec)).toBeNull();
    }
    rec.set(IDENTITY.seen, [7, 3]);
    const seen = seenBy(rec)!;
    expect(seen(Y.createID(7, 2))).toBe(true);
    expect(seen(Y.createID(7, 3))).toBe(false);
    expect(seen(Y.createID(8, 0))).toBe(false);
  });

  test('a wrong "Follow" mark is ignored', () => {
    const doc = new Y.Doc();
    const rec = doc.getMap<unknown>('r');
    rec.set(`${IDENTITY.followed}bad`, [1, 2, 3]);
    rec.set(`${IDENTITY.followed}text`, 'x');
    rec.set(`${IDENTITY.followed}zero`, [5, 0, 7, 3, 0]);
    expect(followsOf(rec)).toEqual([]);
    rec.set(`${IDENTITY.followed}5:10`, [5, 10, 7, 3, 2]);
    const [mark] = followsOf(rec);
    expect(mark).toMatchObject({ client: 5, clock: 10, length: 2 });
    expect(mark!.source(1)).toEqual(Y.createID(7, 4));
  });
});
