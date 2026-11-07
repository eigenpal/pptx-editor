/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Limits of the Yjs session outside the text reads: a room format that arrives after attach,
// the cap on restore marks, and the runs that followed an Enter (`undo-followed.ts`).

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { F, META_KEY, roomOf, SCHEMA_VERSION } from '../collaboration/yjs/schema.ts';
import { RESTORED_KEY } from '../collaboration/yjs/text-identity.ts';
import { MAX_RESTORED } from '../collaboration/yjs/text-identity.ts';
import { restoredTextsOf } from '../collaboration/yjs/undo-restore.ts';
import { checkFollowed, lists, tieFollowedRuns } from '../collaboration/yjs/undo-followed.ts';
import type { IdSet } from '../collaboration/yjs/yjs-items.ts';
import { IN, openPeer, paragraphs, selectIn, sessionsOf } from './lineage-support.ts';

const PART = '/ppt/slides/slide1.xml';

describe('a room format that arrives after attach', () => {
  test('refuses the room, and the editor keeps its own presentation', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Own')) }),
        },
      ],
    });
    const doc = new Y.Doc();
    doc.clientID = 1_001;
    const editor = await openPeer(doc, 'p0', bytes);
    const session = sessionsOf([editor])[0]!;
    // A peer of another format writes the room format after this replica seeded the room.
    const other = new Y.Doc();
    other.clientID = 2_002;
    Y.applyUpdate(other, Y.encodeStateAsUpdate(doc));
    other.getMap(META_KEY).set('schema', SCHEMA_VERSION + 1);
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(other, Y.encodeStateVector(doc)));
    expect(session.statusSnapshot()).toMatchObject({
      status: 'error',
      reason: { code: 'schema-version-mismatch' },
    });
    expect(session.attached).toBe(false);
    // The editor works on its own presentation and writes nothing more to the room.
    const before = Y.encodeStateVector(doc);
    selectIn(editor, 0, 3);
    expect(editor.input('text', 'x').ok).toBe(true);
    expect(paragraphs(editor)).toEqual(['Ownx']);
    expect(Y.encodeStateVector(doc)).toEqual(before);
  });
});

describe('restore marks past the cap', () => {
  test('replicas whose marks arrived in different orders look at the same marks', () => {
    const texts = (junkFirst: boolean): number => {
      const doc = new Y.Doc();
      doc.clientID = 5;
      const room = roomOf(doc);
      const marks = doc.getMap(RESTORED_KEY);
      const text = new Y.Text();
      doc.transact(() => {
        const rec = new Y.Map<unknown>();
        room.nodes.set('k', rec);
        rec.set(F.part, PART);
        rec.set(F.text, text);
        text.insert(0, 'abc');
      });
      const id = text._start!.id;
      const real = () => marks.set(`${id.client}:${id.clock}`, [1, 1, 3]);
      const junk = () => {
        for (let i = 0; i < MAX_RESTORED; i += 1) marks.set(`9:${i}`, [1, 1, 1]);
      };
      doc.transact(() => {
        if (junkFirst) junk();
        real();
        if (!junkFirst) junk();
      });
      return restoredTextsOf(room, new Set([PART])).size;
    };
    expect(texts(true)).toBe(1);
    expect(texts(false)).toBe(1);
  });
});

describe('runs that followed an Enter', () => {
  /** Lists `from` and `to` with run `r` moved from one to the other, as `followEnter` does. */
  function moved() {
    const doc = new Y.Doc();
    doc.clientID = 7;
    const map = doc.getMap('lists');
    const from = new Y.Array<string>();
    const to = new Y.Array<string>();
    const copy = new Y.Text();
    doc.transact(() => {
      map.set('from', from);
      map.set('to', to);
      map.set('copy', copy);
      from.push(['r']);
      to.push(['own']);
      copy.insert(0, 'cd');
    });
    const gone = from._start!.id;
    // The repair keeps the deleted entry, so an undo can put it back (`listEntryId`).
    from._start!.keep = true;
    doc.transact(() => {
      from.delete(0, 1);
      to.push(['r']);
    });
    const added = to._start!.right!.id;
    const anchor = copy._start!.id;
    const set = (id: Y.ID): IdSet => {
      const s = Y.createDeleteSet();
      s.clients.set(id.client, [{ clock: id.clock, len: 1 }] as never);
      return s;
    };
    const step = { insertions: set(anchor), deletions: Y.createDeleteSet(), meta: new Map() };
    tieFollowedRuns([step], [{ anchor, gone, added, run: 'r', from, to }]);
    return { doc, to, step, gone };
  }

  test('an undo puts the old entry back while the new paragraph lists the run', () => {
    const { step, gone } = moved();
    expect([...checkFollowed(step)]).toEqual(['r']);
    expect(Y.isDeleted(step.deletions, gone)).toBe(true);
  });

  test('an undo leaves a run that a peer moved on where it went', () => {
    const { doc, to, step, gone } = moved();
    doc.transact(() => to.delete(1, 1));
    expect([...checkFollowed(step)]).toEqual([]);
    expect(Y.isDeleted(step.deletions, gone)).toBe(false);
  });

  test('a list longer than the walk counts as listing the run', () => {
    const doc = new Y.Doc();
    const list = doc.getArray<string>('long');
    // Each entry inserted at the start is an item of its own.
    doc.transact(() => {
      for (let i = 0; i <= 1 << 16; i += 1) list.insert(0, [`k${i}`]);
    });
    expect(lists(list, 'missing')).toBe(true);
    expect(lists(doc.getArray<string>('short'), 'missing')).toBe(false);
  });
});
