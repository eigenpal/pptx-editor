/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Edge cases of the text lineage read (openspec change `collaboration-text-lineage`) that a
// review found: a join that takes the winner's pieces, the order of a piece's own pieces, a
// rescue that brings back split pieces, a break at the cut, and the order of rescued nodes.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import type { Element, PptxOp } from '@pptx-editor.dev/core/store';
import { F, NODES_KEY, ROOM_LIMITS } from '../collaboration/yjs/schema.ts';
import { LINEAGE, PIECES_KEY } from '../collaboration/yjs/text-lineage.ts';
import {
  all,
  converged,
  firstRow,
  inlineOf,
  kids,
  mergeOps,
  oneRun,
  paragraphs,
  peers,
  sessionsOf,
  root,
  selectIn,
  slide1,
  tableDeck,
  textOf,
  transact,
  TWO_ORDERS,
  twoParagraphs,
} from './lineage-support.ts';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const order = (ids: readonly number[]) => `client ids ${ids.join(', ')}`;

describe('a join that takes the winner pieces', () => {
  for (const ids of TWO_ORDERS)
    test(`the text between two cuts stays when the loser joins (${order(ids)})`, async () => {
      const { editors, sync } = await peers(twoParagraphs(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 1, 3, 11);
      selectIn(b, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectIn(b, 1, 0);
      expect(b.input('deleteBackward').ok).toBe(true);
      sync.flush();
      expect(converged(...editors)).toEqual(['FirstHello world']);
    });

  for (const ids of [
    [1_001, 2_002, 3_003],
    [3_003, 2_002, 1_001],
    [2_002, 3_003, 1_001],
  ])
    test(`the word stays when a third peer joins (${order(ids)})`, async () => {
      const { editors, sync } = await peers(twoParagraphs(), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 1, 6, 11);
      selectIn(b, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectIn(c, 1, 0);
      expect(c.input('deleteBackward').ok).toBe(true);
      sync.flush();
      expect(converged(...editors)).toEqual(['FirstHello world']);
    });
});

describe('the pieces of a piece', () => {
  for (const ids of TWO_ORDERS)
    test(`two later formats of one piece, the second before the first, show the word once (${order(ids)})`, async () => {
      const { editors, sync } = await peers(oneRun('Hello world'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 6, 11);
      selectIn(b, 0, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectIn(b, 0, 9, 11);
      expect(b.exec({ type: 'toggleUnderline' }).ok).toBe(true);
      selectIn(b, 0, 6, 8);
      expect(b.exec({ type: 'toggleUnderline' }).ok).toBe(true);
      sync.flush();
      expect(converged(...editors)).toEqual(['Hello world']);
    });

  for (const ids of TWO_ORDERS)
    test(`an edited losing piece takes its own pieces with it, in order (${order(ids)})`, async () => {
      const { editors, sync } = await peers(oneRun('Hello world'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 6, 11);
      selectIn(b, 0, 7, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectIn(b, 0, 9, 11);
      expect(b.exec({ type: 'toggleUnderline' }).ok).toBe(true);
      selectIn(b, 0, 9);
      expect(b.input('text', 'X').ok).toBe(true);
      sync.flush();
      // The loser's edited text shows once more, in order, after the winner's word.
      expect(converged(...editors)).toEqual(['Hello worldorXld']);
    });

  for (const ids of TWO_ORDERS)
    test(`an edited losing piece goes after the winner's last piece in document order (${order(ids)})`, async () => {
      const { editors, sync } = await peers(oneRun('Hello world'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 3, 8);
      selectIn(b, 0, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      selectIn(a, 0, 8);
      expect(a.input('paragraph').ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectIn(b, 0, 11);
      expect(b.input('text', '!').ok).toBe(true);
      sync.flush();
      expect(converged(...editors)).toEqual(['Hello wo', 'rldworld!']);
    });
});

describe('a line break at the cut', () => {
  for (const ids of TWO_ORDERS)
    test(`the break stays before the word, not after it (${order(ids)})`, async () => {
      const { editors, sync } = await peers(oneRun('Hello world'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 0, 6, 11);
      selectIn(b, 0, 6);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.input('lineBreak').ok).toBe(true);
      sync.flush();
      expect(converged(...editors)).toEqual(['Hello world']);
      expect(inlineOf(a, 'Hello')).toEqual(['Hello ', '<br>', 'world']);
    });
});

/** The ops of a format of the lower cell's run from `at` on: the run keeps its start. */
function splitOps(e: PptxEditorInstance, at: number, attr: 'b' | 'i'): PptxOp[] {
  const part = slide1(e);
  const lower = all(root(e), 'tc')[1]!;
  const p = all(lower, 'p')[0]!;
  const r = kids(p).find((c) => c.local === 'r')!;
  const t = kids(r).find((c) => c.local === 't')!;
  const text = textOf(r);
  return [
    { op: 'setText', part, node: t.id, text: text.slice(0, at) },
    {
      op: 'insertNode',
      part,
      parent: p.id,
      index: p.children.indexOf(r as Element) + 1,
      spec: {
        ns: A_NS,
        local: 'r',
        prefix: 'a',
        children: [
          {
            ns: A_NS,
            local: 'rPr',
            prefix: 'a',
            attrs: [
              { ns: '', local: 'lang', value: 'en-US' },
              { ns: '', local: attr, value: '1' },
            ],
          },
          { ns: A_NS, local: 't', prefix: 'a', children: [{ text: text.slice(at) }] },
        ],
      },
    },
  ];
}

describe('a rescue of a paragraph that holds split pieces', () => {
  for (const ids of TWO_ORDERS)
    test(`the word shows once after the rescue (${order(ids)})`, async () => {
      const { editors, sync } = await peers(tableDeck('Hello world'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      transact(a, splitOps(a, 6, 'b'));
      transact(b, splitOps(b, 6, 'i'));
      sync.flush();
      expect(converged(...editors)).toEqual(['A', 'Hello world']);
      sync.pause();
      transact(a, mergeOps(a));
      transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
      sync.flush();
      expect(converged(...editors)).toEqual(['Hello world']);
    });
});

describe('the order of rescued nodes', () => {
  for (const ids of TWO_ORDERS)
    test(`nodes moved in two writes from different places come back in their order (${order(ids)})`, async () => {
      const { editors, sync } = await peers(tableDeck('B1\nB2\nB3'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      const part = slide1(a);
      const moveUp = (index: number) => {
        const [top, bottom] = all(root(a), 'tc');
        const topBody = kids(top!).find((c) => c.local === 'txBody')!;
        transact(a, [
          {
            op: 'moveNode',
            part,
            node: all(bottom!, 'p')[index]!.id,
            parent: topBody.id,
            index: topBody.children.length,
          },
        ]);
      };
      sync.pause();
      moveUp(0);
      // B3 is now at index 1 of the lower cell.
      moveUp(1);
      transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
      sync.flush();
      expect(converged(...editors)).toEqual(['B1', 'B2', 'B3']);
    });
});

describe('the piece cap counts live pieces', () => {
  test('many removed pieces of old formats keep the join rule on', async () => {
    const { editors, sync, docs } = await peers(twoParagraphs(), [1_001, 2_002]);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    // Records of pieces that users deleted long ago: removed, but still in the index.
    docs[0]!.transact(() => {
      const index = docs[0]!.getMap(PIECES_KEY);
      const nodes = docs[0]!.getMap('pptx-nodes-v1');
      for (let i = 0; i < 5_000; i += 1) {
        const key = `zz.${i}`;
        const rec = new Y.Map<unknown>();
        rec.set(F.part, slide1(a));
        rec.set(F.key, key);
        rec.set(LINEAGE.splitFrom, 'zz.origin');
        rec.set(LINEAGE.splitBaseText, 'old');
        rec.set(LINEAGE.splitAt, 1);
        rec.set(LINEAGE.splitLength, 2);
        rec.set(LINEAGE.splitGeneration, 1);
        rec.set(`${LINEAGE.removedWith}.zz`, [key]);
        nodes.set(key, rec);
        index.set(key, slide1(a));
      }
    });
    selectIn(a, 1, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    selectIn(b, 1, 0);
    expect(b.input('deleteBackward').ok).toBe(true);
    sync.flush();
    expect(converged(...editors)).toEqual(['FirstHello world']);
  });
});

describe('a rescue read that the room refuses', () => {
  test('reaches each session as remote-apply-failed, keeps the text, and heals', async () => {
    const { editors, sync, docs } = await peers(tableDeck(), [1_001, 2_002]);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    sync.pause();
    transact(a, mergeOps(a));
    transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
    sync.flush();
    expect(converged(...editors)).toEqual(['B']);
    // A hostile peer gives the rescued paragraph more copies of `a:pPr` than a read accepts.
    // No list under the slide holds that paragraph, so only the rescue reads it.
    const doc = docs[0]!;
    const nodes = doc.getMap(NODES_KEY) as Y.Map<Y.Map<unknown>>;
    const textUnder = (key: string, depth = 0): string => {
      const rec = nodes.get(key);
      const t = rec?.get(F.text);
      if (t instanceof Y.Text) return t.toString();
      const list = rec?.get(F.children);
      return list instanceof Y.Array && depth < 8
        ? (list.toArray() as string[]).map((k) => textUnder(k, depth + 1)).join('')
        : '';
    };
    const pKey = [...nodes.keys()].find((k) => {
      const shell = nodes.get(k)!.get(F.shell) as { local?: string } | undefined;
      return shell?.local === 'p' && textUnder(k) === 'B';
    })!;
    const pRec = nodes.get(pKey)!;
    const hostile = Array.from({ length: ROOM_LIMITS.maxCopiesPerGroup + 2 }, (_, i) => `99.h${i}`);
    doc.transact(() => {
      for (const k of hostile) {
        const rec = new Y.Map<unknown>();
        rec.set(F.part, slide1(a));
        rec.set(F.key, k);
        rec.set(F.shell, { ns: A_NS, local: 'pPr', prefix: 'a', nsDecls: [] });
        rec.set(F.attrs, new Y.Map());
        rec.set(F.children, new Y.Array());
        nodes.set(k, rec);
      }
      (pRec.get(F.children) as Y.Array<string>).insert(0, hostile);
    }, 'hostile');
    for (const s of sessionsOf(editors))
      expect(s.statusSnapshot()).toMatchObject({
        status: 'error',
        reason: { code: 'remote-apply-failed' },
      });
    // Each store keeps its last good state.
    expect(paragraphs(a)).toEqual(['B']);
    expect(paragraphs(b)).toEqual(['B']);
    // The copies go, and every replica reads the part again.
    doc.transact(() => {
      const list = pRec.get(F.children) as Y.Array<string>;
      list.delete((list.toArray() as string[]).indexOf(hostile[0]!), hostile.length);
      for (const k of hostile) nodes.delete(k);
    }, 'hostile');
    expect(converged(...editors)).toEqual(['B']);
  });
});
