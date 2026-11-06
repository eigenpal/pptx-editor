/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The rules of `text-lineage.ts` on a room that a hostile peer wrote: each read stays linear in
// what it looks at, and a cap or a guard keeps every piece instead of dropping text.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { ReplicaElementSpec } from '@pptx-editor.dev/core/store';
import { roomOf } from '../collaboration/yjs/schema.ts';
import {
  LINEAGE,
  MOVES_KEY,
  PIECES_KEY,
  SPLITS_KEY,
  dropLosingSplits,
  placeEditedPieces,
  splitIndexKey,
} from '../collaboration/yjs/text-lineage.ts';
import { rescueMoved } from '../collaboration/yjs/move-rescue.ts';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART = '/ppt/slides/slide1.xml';
const P = { ns: A_NS, local: 'p' };

const runSpec = (key: string, text: string): ReplicaElementSpec => ({
  key,
  ns: A_NS,
  local: 'r',
  children: [{ key: `${key}t`, ns: A_NS, local: 't', children: [{ key: `${key}x`, text }] }],
});

/**
 * A chain of `n` pieces in one paragraph: each one split from the one before it, by two
 * clients at once, so every group has a winner and every piece is looked at.
 */
function chain(n: number) {
  const doc = new Y.Doc();
  const room = roomOf(doc);
  const children: ReplicaElementSpec[] = [];
  doc.transact(() => {
    const splits = doc.getMap(SPLITS_KEY);
    const pieces = doc.getMap(PIECES_KEY);
    let from = '0.root';
    for (let i = 0; i < n; i += 1) {
      const key = `${i % 2 === 0 ? 'b' : 'a'}.${i}`;
      const rec = new Y.Map<unknown>();
      room.nodes.set(key, rec);
      rec.set('p', PART);
      rec.set(LINEAGE.splitFrom, from);
      rec.set(LINEAGE.splitBaseText, 'xy');
      rec.set(LINEAGE.splitAt, 1);
      rec.set(LINEAGE.splitLength, 1);
      rec.set(LINEAGE.splitGeneration, 1);
      pieces.set(key, PART);
      const group = splitIndexKey(from, 'xy', 1);
      splits.set(`${group}\u0000a`, 1);
      splits.set(`${group}\u0000b`, 1);
      children.push(runSpec(key, 'y'));
      from = key;
    }
  });
  return { room, children };
}

describe('bounded work on a hostile room', () => {
  for (const n of [4_000, 20_000])
    test(`a chain of ${n} split pieces reads in linear time`, () => {
      const { room, children } = chain(n);
      const started = performance.now();
      const shown = dropLosingSplits(room, P, children, new Set());
      const spec: ReplicaElementSpec = { key: 'p', ...P, children: [...shown] };
      placeEditedPieces(room, PART, spec, () => null);
      expect(performance.now() - started).toBeLessThan(1_000);
      // Over the cap, every piece stays: a cap never drops text.
      if (n > 4_096) expect(shown).toBe(children);
    });

  test('moved keys against a removal snapshot of many keys read in linear time', () => {
    const doc = new Y.Doc();
    const room = roomOf(doc);
    const snapshot = Array.from({ length: 20_000 }, (_, i) => `z.${i}`);
    doc.transact(() => {
      const anc = new Y.Map<unknown>();
      room.nodes.set('r.1', anc);
      anc.set(LINEAGE.removedWith, snapshot);
      for (let i = 0; i < 3_000; i += 1) {
        const rec = new Y.Map<unknown>();
        room.nodes.set(`m.${i}`, rec);
        rec.set('p', PART);
        rec.set(LINEAGE.movedFrom, 'p');
        rec.set(LINEAGE.movedTo, ['r.1']);
        doc.getMap(MOVES_KEY).set(`m.${i}`, PART);
      }
    });
    const spec: ReplicaElementSpec = { key: 'p', ...P, children: [] };
    const started = performance.now();
    let reads = 0;
    rescueMoved(room, PART, spec, () => {
      reads += 1;
      return null;
    });
    expect(performance.now() - started).toBeLessThan(1_000);
    // Over the rescue cap: none is put back, and none is read.
    expect(reads).toBe(0);
  });
});

describe('hostile values are ignored', () => {
  /** One run split once by two clients; client `a` cut it first. */
  function contest() {
    const doc = new Y.Doc();
    const room = roomOf(doc);
    const group = splitIndexKey('0.root', 'Hello world', 1);
    doc.transact(() => {
      for (const client of ['a', 'b']) {
        const rec = new Y.Map<unknown>();
        room.nodes.set(`${client}.1`, rec);
        rec.set('p', PART);
        rec.set(LINEAGE.splitFrom, '0.root');
        rec.set(LINEAGE.splitBaseText, 'Hello world');
        rec.set(LINEAGE.splitAt, 6);
        rec.set(LINEAGE.splitLength, 5);
        rec.set(LINEAGE.splitGeneration, 1);
      }
      doc.getMap(SPLITS_KEY).set(`${group}\u0000a`, 6);
      doc.getMap(SPLITS_KEY).set(`${group}\u0000b`, 6);
    });
    const children = [
      runSpec('0.root', 'Hello '),
      runSpec('a.1', 'world'),
      runSpec('b.1', 'world'),
    ];
    return { doc, room, group, children };
  }
  const keys = (specs: readonly { key: string }[]) => specs.map((s) => s.key);

  test('a fair contest keeps the winner only', () => {
    const { room, children } = contest();
    expect(keys(dropLosingSplits(room, P, children, new Set()))).toEqual(['0.root', 'a.1']);
  });

  for (const [name, cut] of [
    ['NaN', Number.NaN],
    ['negative', -5],
    ['infinite', Number.NEGATIVE_INFINITY],
    ['fractional', 0.5],
  ] as const)
    test(`a ${name} cut in the split index wins nothing`, () => {
      const { doc, room, group, children } = contest();
      doc.getMap(SPLITS_KEY).set(`${group}\u00000`, cut);
      expect(keys(dropLosingSplits(room, P, children, new Set()))).toEqual(['0.root', 'a.1']);
    });

  test('a generation of 2^31 or more makes no piece, and the text stays once', () => {
    const { room, children } = contest();
    (room.nodes.get('a.1') as Y.Map<unknown>).set(LINEAGE.splitGeneration, 2 ** 31);
    // `a.1` reads as a plain run and stays; `b.1` still loses to client a's split.
    expect(keys(dropLosingSplits(room, P, children, new Set()))).toEqual(['0.root', 'a.1']);
  });

  for (const [name, field, value] of [
    ['a negative cut', LINEAGE.splitAt, -1],
    ['a NaN length', LINEAGE.splitLength, Number.NaN],
    ['a string base text', LINEAGE.splitBaseText, 7],
  ] as const)
    test(`${name} on a piece leaves it as listed`, () => {
      const { room, children } = contest();
      (room.nodes.get('b.1') as Y.Map<unknown>).set(field, value);
      expect(keys(dropLosingSplits(room, P, children, new Set()))).toEqual(keys(children));
    });

  for (const [name, set] of [
    ['a movedTo that is no array', (r: Y.Map<unknown>) => r.set(LINEAGE.movedTo, 'r.1')],
    ['a movedFrom that is no string', (r: Y.Map<unknown>) => r.set(LINEAGE.movedFrom, 5)],
    [
      'a movedTo longer than any writer makes',
      (r: Y.Map<unknown>) => r.set(LINEAGE.movedTo, Array(400).fill('r.1')),
    ],
    [
      'a removal snapshot over the cap',
      (r: Y.Map<unknown>, room: ReturnType<typeof roomOf>) =>
        (room.nodes.get('r.1') as Y.Map<unknown>).set(LINEAGE.removedWith, Array(20_002).fill('z')),
    ],
  ] as const)
    test(`${name} rescues nothing`, () => {
      const doc = new Y.Doc();
      const room = roomOf(doc);
      const rec = new Y.Map<unknown>();
      doc.transact(() => {
        const anc = new Y.Map<unknown>();
        room.nodes.set('r.1', anc);
        anc.set(LINEAGE.removedWith, []);
        room.nodes.set('m.1', rec);
        rec.set('p', PART);
        rec.set(LINEAGE.movedFrom, 'p');
        rec.set(LINEAGE.movedTo, ['r.1']);
        doc.getMap(MOVES_KEY).set('m.1', PART);
      });
      const spec: ReplicaElementSpec = { key: 'p', ...P, children: [] };
      const node: ReplicaElementSpec = { key: 'm.1', ...P, children: [] };
      // Without the hostile value the node comes back.
      expect(rescueMoved(room, PART, spec, () => node).children).toHaveLength(1);
      set(rec, room);
      expect(rescueMoved(room, PART, spec, () => node).children).toHaveLength(0);
    });
});

describe('a moved node whose old parents are all gone', () => {
  /**
   * Runs `m.1` and `m.2` moved from the gone paragraph `p.3` into `r.2` (a join); a peer then
   * moved `c.1` out of `r.2` into `p.1` and removed `r.2` without having seen them (its own join).
   */
  function joined(shown: boolean) {
    const doc = new Y.Doc();
    const room = roomOf(doc);
    doc.transact(() => {
      const removed = new Y.Map<unknown>();
      room.nodes.set('r.2', removed);
      removed.set(LINEAGE.removedWith, ['r.2', 'c.1']);
      const kept = new Y.Map<unknown>();
      room.nodes.set('c.1', kept);
      kept.set('p', PART);
      kept.set(LINEAGE.movedFrom, 'r.2');
      kept.set(LINEAGE.movedTo, ['p.1']);
      doc.getMap(MOVES_KEY).set('c.1', PART);
      // Listed out of order, so the read sorts them by the index they left.
      for (const [key, index] of [
        ['m.2', 1],
        ['m.1', 0],
      ] as const) {
        const rec = new Y.Map<unknown>();
        room.nodes.set(key, rec);
        rec.set('p', PART);
        rec.set(LINEAGE.movedFrom, 'p.3');
        rec.set(LINEAGE.movedIndex, index);
        rec.set(LINEAGE.movedClock, 7);
        rec.set(LINEAGE.movedTo, ['r.2', 'body']);
        doc.getMap(MOVES_KEY).set(key, PART);
      }
    });
    const end: ReplicaElementSpec = { key: 'e', ns: A_NS, local: 'endParaRPr', children: [] };
    const first: ReplicaElementSpec = {
      key: 'p.1',
      ...P,
      children: [runSpec('a.1', 'AB'), ...(shown ? [runSpec('c.1', 'CD')] : []), end],
    };
    const spec: ReplicaElementSpec = { key: 'body', ns: A_NS, local: 'txBody', children: [first] };
    return { room, spec };
  }
  const read = (key: string) => runSpec(key, key);
  const keysOf = (spec: ReplicaElementSpec) =>
    (spec.children![0] as ReplicaElementSpec).children!.map((c) => c.key);

  test('follows the children that the removal moved out, in the order it left', () => {
    const { room, spec } = joined(true);
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['a.1', 'c.1', 'm.1', 'm.2', 'e']);
  });

  test('stays out when no child that the removal moved out shows', () => {
    const { room, spec } = joined(false);
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['a.1', 'e']);
  });

  test('stays out when the removal saw it', () => {
    const { room, spec } = joined(true);
    (room.nodes.get('r.2') as Y.Map<unknown>).set(LINEAGE.removedWith, ['r.2', 'c.1', 'm.1']);
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['a.1', 'c.1', 'm.2', 'e']);
  });
});

describe('nodes that follow a join on a hostile room', () => {
  /**
   * Orphans `o.<i>`, each moved out of `x.<i>` into `x.<i - 1>`, which a removal that did not
   * see it took: each one follows the one before it, and `o.0` follows `c.1`, a run moved out
   * of `x.-1`, when `anchored`. `links` changes the parent that an orphan went into.
   */
  function orphans(n: number, anchored: boolean, links: Readonly<Record<number, number>> = {}) {
    const doc = new Y.Doc();
    const room = roomOf(doc);
    doc.transact(() => {
      const moves = doc.getMap(MOVES_KEY);
      const anchor = new Y.Map<unknown>();
      room.nodes.set('c.1', anchor);
      anchor.set('p', PART);
      anchor.set(LINEAGE.movedFrom, 'x.-1');
      anchor.set(LINEAGE.movedTo, ['p.1']);
      moves.set('c.1', PART);
      for (let i = 0; i < n; i += 1) {
        const into = `x.${links[i] ?? i - 1}`;
        if (!room.nodes.has(into)) {
          const gone = new Y.Map<unknown>();
          room.nodes.set(into, gone);
          gone.set(LINEAGE.removedWith, [into]);
        }
        const rec = new Y.Map<unknown>();
        room.nodes.set(`o.${i}`, rec);
        rec.set('p', PART);
        rec.set(LINEAGE.movedFrom, `x.${i}`);
        rec.set(LINEAGE.movedIndex, 0);
        rec.set(LINEAGE.movedClock, 1);
        rec.set(LINEAGE.movedTo, [into, 'body']);
        moves.set(`o.${i}`, PART);
      }
    });
    const first: ReplicaElementSpec = {
      key: 'p.1',
      ...P,
      children: [runSpec('a.1', 'AB'), ...(anchored ? [runSpec('c.1', 'CD')] : [])],
    };
    const spec: ReplicaElementSpec = { key: 'body', ns: A_NS, local: 'txBody', children: [first] };
    return { room, spec };
  }
  const read = (key: string) => runSpec(key, key);
  const keysOf = (spec: ReplicaElementSpec) =>
    (spec.children![0] as ReplicaElementSpec).children!.map((c) => c.key);

  test('a cycle of anchors with nothing shown places nothing', () => {
    // `o.0` went into `x.1`, and `o.1` into `x.0`: each waits for the other.
    const { room, spec } = orphans(2, true, { 0: 1 });
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['a.1', 'c.1']);
  });

  test('a cycle of anchors places only what follows a shown node', () => {
    // `o.0` follows `c.1`; `o.1` went into `x.2` and `o.2` into `x.1`, so they stay out.
    const { room, spec } = orphans(3, true, { 1: 2 });
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['a.1', 'c.1', 'o.0']);
  });

  test('a chain of 1024 orphans, one placed in each round, reads in bounded time', () => {
    const { room, spec } = orphans(1_024, true);
    const started = performance.now();
    const keys = keysOf(rescueMoved(room, PART, spec, read));
    expect(performance.now() - started).toBeLessThan(1_000);
    expect(keys).toEqual(['a.1', 'c.1', ...Array.from({ length: 1_024 }, (_, i) => `o.${i}`)]);
  });

  test('one orphan over the rescue cap puts none back', () => {
    const { room, spec } = orphans(1_025, true);
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['a.1', 'c.1']);
  });

  test('prototype names are plain keys', () => {
    const doc = new Y.Doc();
    const room = roomOf(doc);
    doc.transact(() => {
      const gone = new Y.Map<unknown>();
      room.nodes.set('__proto__', gone);
      gone.set(LINEAGE.removedWith, ['__proto__']);
      for (const [key, from, into] of [
        ['toString', '__proto__', 'p.1'],
        ['hasOwnProperty', 'constructor', '__proto__'],
      ] as const) {
        const rec = new Y.Map<unknown>();
        room.nodes.set(key, rec);
        rec.set('p', PART);
        rec.set(LINEAGE.movedFrom, from);
        rec.set(LINEAGE.movedTo, [into, 'body']);
        doc.getMap(MOVES_KEY).set(key, PART);
      }
    });
    const first: ReplicaElementSpec = { key: 'p.1', ...P, children: [runSpec('toString', 'CD')] };
    const spec: ReplicaElementSpec = { key: 'body', ns: A_NS, local: 'txBody', children: [first] };
    expect(keysOf(rescueMoved(room, PART, spec, read))).toEqual(['toString', 'hasOwnProperty']);
  });
});
