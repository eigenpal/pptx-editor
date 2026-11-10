/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Table rows deleted while peers move or merge their text (AGENTS.md, "Collaboration
// verification"): editors on linked Y.Docs, the link paused, both client orders, undo and redo,
// and a delete while a peer types. Every case checks that the replicas hold the same package,
// that every session applied every remote change, and that the saved deck reopens with the
// same text and passes the part checks.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  textOf,
  type Element,
  type PptxOp,
} from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { F, roomOf } from '../collaboration/yjs/schema.ts';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import {
  IN,
  TWO_ORDERS,
  all,
  converged,
  deleteMergedFirstRowOps,
  kids,
  mergeOps,
  paragraphs,
  peers,
  root,
  sessionsOf,
  slide1,
  tableDeck,
  transact,
} from './lineage-support.ts';

/** The saved deck of `e` reopens, and no part check fails on it. */
function savedValid(e: PptxEditorInstance): void {
  const r = readPptx(e.save()!);
  expect(r.ok).toBe(true);
  if (r.ok) expect(packageViolations(r.pkg)).toEqual([]);
}

/** Two replicas: the lower cell holds "PQRS" and "TUVW". */
async function twoPeers(order: readonly number[]) {
  const { editors, sync, docs } = await peers(tableDeck('PQRS\nTUVW'), order);
  return { a: editors[0]!, b: editors[1]!, sync, docs };
}

/** The op that moves the second paragraph of the lower cell before the first. */
function moveSecondFirst(e: PptxEditorInstance): PptxOp {
  const body = kids(all(root(e), 'tc')[1]!).find((c) => c.local === 'txBody')!;
  const [first, second] = kids(body).filter((c) => c.local === 'p');
  return {
    op: 'moveNode',
    part: slide1(e),
    node: second!.id,
    parent: body.id,
    index: body.children.indexOf(first!),
  };
}

/**
 * Peer A moves a paragraph of the lower cell while peer B merges the two cells, so both move
 * one paragraph at once: the room lists it in both cells, and a read shows it once.
 */
async function concurrentMoves(order: readonly number[]) {
  const peersOf = await twoPeers(order);
  const { a, b, sync } = peersOf;
  sync.pause();
  transact(a, [moveSecondFirst(a)]);
  transact(b, mergeOps(b));
  sync.flush();
  expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW', '']);
  return peersOf;
}

const orderName = (order: readonly number[]) => (order[0]! < order[1]! ? 'A, B' : 'B, A');

describe('a row delete after two peers moved one paragraph at once', () => {
  for (const order of TWO_ORDERS)
    test(`both replicas remove the paragraph with its row (client order ${orderName(order)})`, async () => {
      const { a, b } = await concurrentMoves(order);
      transact(a, deleteMergedFirstRowOps(a));
      expect(converged(a, b)).toEqual(['']);
      savedValid(a);
    });

  for (const order of TWO_ORDERS)
    test(`undo and redo of the row delete reach both replicas (client order ${orderName(order)})`, async () => {
      const { a, b } = await concurrentMoves(order);
      transact(a, deleteMergedFirstRowOps(a));
      expect(converged(a, b)).toEqual(['']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW', '']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['']);
      savedValid(b);
    });

  for (const order of TWO_ORDERS)
    test(`a peer that types in the paragraph while the row goes keeps its typing (client order ${orderName(order)})`, async () => {
      const { a, b, sync } = await concurrentMoves(order);
      sync.pause();
      const t = all(root(b), 't').find((x) => textOf(x) === 'PQRS')!;
      transact(b, [{ op: 'setText', part: slide1(b), node: t.id, text: 'PQRS!' }]);
      transact(a, deleteMergedFirstRowOps(a));
      sync.flush();
      expect(converged(a, b)).toEqual(['!']);
      // The typist's undo reverts its typing only: the paragraph does not come back.
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['']);
      savedValid(a);
    });

  for (const order of TWO_ORDERS)
    test(`a local move of the paragraph into the other cell lists it there once (client order ${orderName(order)})`, async () => {
      const { a, b } = await concurrentMoves(order);
      const pqrs = all(root(a), 'p').find((p) => textOf(p) === 'PQRS')!;
      const body = kids(all(root(a), 'tc')[1]!).find((c) => c.local === 'txBody')!;
      const firstP = kids(body).find((c) => c.local === 'p')!;
      transact(a, [
        {
          op: 'moveNode',
          part: slide1(a),
          node: pqrs.id,
          parent: body.id,
          index: body.children.indexOf(firstP),
        },
      ]);
      expect(converged(a, b)).toEqual(['A', 'TUVW', 'PQRS']);
      transact(b, deleteMergedFirstRowOps(b));
      expect(converged(a, b)).toEqual(['PQRS']);
      savedValid(a);
    });
});

/** The child lists of the room on `doc` that list the paragraph whose text is `text`. */
function listingsOf(doc: Y.Doc, text: string): { paragraph: string; lists: Y.Array<string>[] } {
  const records = [...roomOf(doc).nodes.entries()].flatMap(([k, r]) =>
    r instanceof Y.Map ? [[k, r] as const] : []
  );
  const listing = (key: string) =>
    records.flatMap(([, r]) => {
      const list = r.get(F.children);
      return list instanceof Y.Array && list.toArray().includes(key) ? [list] : [];
    });
  const keyOf = (list: Y.Array<string> | undefined) =>
    records.find(([, r]) => r.get(F.children) === list)?.[0];
  const textKey = records.find(([, r]) => {
    const x = r.get(F.text);
    return x instanceof Y.Text && x.toString() === text;
  })?.[0];
  if (textKey === undefined) throw new Error(`no text record "${text}" in the room`);
  // The text node's record → its a:t → its a:r → its a:p.
  const paragraph = keyOf(listing(keyOf(listing(keyOf(listing(textKey)[0])!)[0])!)[0]);
  if (paragraph === undefined) throw new Error(`no paragraph lists "${text}"`);
  return { paragraph, lists: listing(paragraph) };
}

/** The op that writes `value` into `bandRow` of the table: a change off every cell's path. */
const bandRow = (e: PptxEditorInstance, value: string): PptxOp => ({
  op: 'setAttribute',
  part: slide1(e),
  node: all(root(e), 'tblPr')[0]!.id,
  name: { ns: '', local: 'bandRow' },
  value,
});

/** The saved deck of `e`, reopened, as a fingerprint. */
function savedFingerprint(e: PptxEditorInstance): string {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error(`the saved deck does not reopen: ${r.error.reason}`);
  return canonicalPackageFingerprint(r.pkg);
}

describe('the listing of a paragraph that two peers moved at once and the read skips', () => {
  for (const order of TWO_ORDERS)
    test(`both replicas save the same deck before any further edit (client order ${orderName(order)})`, async () => {
      const { a, b, docs } = await concurrentMoves(order);
      expect(listingsOf(docs[0]!, 'PQRS').lists).toHaveLength(2);
      expect(savedFingerprint(b)).toBe(savedFingerprint(a));
    });

  for (const order of TWO_ORDERS)
    test(`a later edit deletes it, and no undo of either move brings it back (client order ${orderName(order)})`, async () => {
      const { a, b, docs } = await concurrentMoves(order);
      transact(a, [bandRow(a, '1')]);
      expect(listingsOf(docs[1]!, 'PQRS').lists).toHaveLength(1);
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW', '']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW', '']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW']);
      for (const doc of docs) expect(listingsOf(doc, 'PQRS').lists).toHaveLength(1);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW', '']);
      expect(savedFingerprint(b)).toBe(savedFingerprint(a));
    });

  for (const order of TWO_ORDERS)
    test(`a remote change that the replica refuses leaves every listing in place (client order ${orderName(order)})`, async () => {
      const { a, b, docs } = await concurrentMoves(order);
      // A peer takes the shown listing away, in a transaction with a part entry that every
      // replica refuses, so the rebuild stops before it reads the slide again.
      const room = roomOf(docs[1]!);
      const { paragraph, lists } = listingsOf(docs[1]!, 'PQRS');
      const shown = lists.find((l) => l.length === Math.max(...lists.map((x) => x.length)))!;
      docs[1]!.transact(() => {
        room.parts.set('/a-refused.xml', { kind: 'xml', root: 'no-record' });
        shown.delete(shown.toArray().indexOf(paragraph), 1);
      });
      expect(sessionsOf([a])[0]!.statusSnapshot()).toMatchObject({ status: 'error' });
      // An edit on the replica that refused: its notes are of an older read, so they delete
      // nothing, and the paragraph keeps the listing that now shows it.
      transact(a, [bandRow(a, '1')]);
      expect(listingsOf(docs[1]!, 'PQRS').lists).toHaveLength(1);
      docs[1]!.transact(() => room.parts.delete('/a-refused.xml'));
      // The listing left is in the covered cell, which no writer saw it in: it shows in the
      // merged cell (`merged-cells.ts`).
      expect(paragraphs(b)).toEqual(['A', 'TUVW', 'PQRS', '']);
    });

  for (const order of TWO_ORDERS)
    test(`a room that lists one paragraph many times in one cell: an edit deletes only the repeats (client order ${orderName(order)})`, async () => {
      const { a, b, docs } = await twoPeers(order);
      const { paragraph, lists } = listingsOf(docs[1]!, 'PQRS');
      const list = lists[0]!;
      // A hostile peer repeats the paragraph after its own listing and at the end of the list.
      const at = list.toArray().indexOf(paragraph);
      docs[1]!.transact(() => {
        list.insert(
          at + 1,
          Array.from({ length: 500 }, () => paragraph)
        );
        list.push(Array.from({ length: 500 }, () => paragraph));
      });
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW']);
      transact(a, [bandRow(a, '1')]);
      expect(list.toArray().filter((k) => k === paragraph)).toEqual([paragraph]);
      expect(list.toArray().indexOf(paragraph)).toBe(at);
      expect(converged(a, b)).toEqual(['A', 'PQRS', 'TUVW']);
    });
});

// A vertical merge moves the lower cell's text into the upper cell. When peers delete both rows
// at once, the text goes with them: the peer that deleted the lower row saw the text there, and
// every order of the three edits ends without it (the merge before both deletes, or a delete
// before the merge, which then has no cell to merge). When the lower row's delete came after the
// merge, the delete of the upper row did not see the text, but the text has no cell left to go
// back to, so it goes with the merged cell.

const IN_ROW = 370_840;
/** A three-row table: "A|X", "B|Y", "C|Z". */
const threeRows = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: 3 * IN_ROW,
          columns: [IN, IN],
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: ['AX', 'BY', 'CZ']
            .map((r) => `<a:tr h="${IN_ROW}">${[...r].map((c) => cell(c)).join('')}</a:tr>`)
            .join(''),
        }),
      },
    ],
  });

/** A three-row table of one column: "A", "B", "C". */
const threeRowsOneColumn = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: IN,
          cy: 3 * IN_ROW,
          columns: [IN],
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: [...'ABC'].map((c) => `<a:tr h="${IN_ROW}">${cell(c)}</a:tr>`).join(''),
        }),
      },
    ],
  });

const THREE_ORDERS: readonly (readonly number[])[] = [
  [1_001, 2_002, 3_003],
  [3_003, 2_002, 1_001],
  [2_002, 3_003, 1_001],
];

/** The graphic frames under `el`. */
const framesIn = (el: Element): Element[] =>
  el.kind === 'graphicFrame' ? [el] : kids(el).flatMap(framesIn);
const frame = (e: PptxEditorInstance): Element => framesIn(root(e))[0]!;

/** The text of every cell, row by row. */
const grid = (e: PptxEditorInstance): string[] =>
  all(root(e), 'tr').map((tr) =>
    kids(tr)
      .filter((tc) => tc.local === 'tc')
      .map((tc) => textOf(tc))
      .join('|')
  );

const caret = (e: PptxEditorInstance, row: number, column: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: frame(e).id,
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });

/** Merge the first two cells of the first column. */
function mergeFirstColumn(e: PptxEditorInstance): void {
  e.setSelection({
    kind: 'cells',
    nodeId: frame(e).id,
    anchor: { row: 0, column: 0 },
    focus: { row: 1, column: 0 },
  });
  expect(e.exec({ type: 'mergeCells' }).ok).toBe(true);
}

function deleteRow(e: PptxEditorInstance, row: number): void {
  caret(e, row, 1);
  expect(e.exec({ type: 'deleteRow' }).ok).toBe(true);
}

/** Every session applied every remote change, the replicas match, and each save is valid. */
function tablesConverged(editors: readonly PptxEditorInstance[]): string[] {
  for (const s of sessionsOf(editors))
    expect(s.statusSnapshot()).toMatchObject({ status: 'ready' });
  const [first, ...rest] = editors;
  for (const e of rest)
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(
      canonicalPackageFingerprint(first!.session!.store.package)
    );
  for (const e of editors) savedValid(e);
  return grid(first!);
}

const names = (order: readonly number[]) => order.map((n) => n / 1_001).join(', ');

describe('a merge while peers delete both rows of the merged cell', () => {
  for (const order of THREE_ORDERS)
    test(`the merged text goes with the rows (client order ${names(order)})`, async () => {
      const { editors, sync } = await peers(threeRows(), order);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a);
      deleteRow(b, 1);
      deleteRow(c, 0);
      sync.flush();
      expect(tablesConverged(editors)).toEqual(['C|Z']);
    });

  for (const order of THREE_ORDERS)
    test(`the merge as store ops: the moved paragraph goes with the rows (client order ${names(order)})`, async () => {
      const { editors, sync } = await peers(threeRowsOneColumn(), order);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      transact(a, mergeOps(a));
      transact(b, [{ op: 'removeNode', part: slide1(b), node: all(root(b), 'tr')[1]!.id }]);
      transact(c, [{ op: 'removeNode', part: slide1(c), node: all(root(c), 'tr')[0]!.id }]);
      sync.flush();
      expect(converged(...editors)).toEqual(['C']);
      savedValid(a);
    });

  for (const order of TWO_ORDERS)
    test(`the delete of the lower row after the merge, and a delete of the upper row that did not see it keep the text (client order ${names(order)})`, async () => {
      const { editors, sync, docs } = await peers(threeRows(), [...order, 3_003]);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a);
      sync.deliverTo(docs[1]!);
      expect(grid(b)[0]!.split('|')[0]).toBe('AB');
      deleteRow(b, 1);
      deleteRow(c, 0);
      sync.flush();
      // B saw "B" leave its row, and C did not see "B" in the merged cell: no peer removed it,
      // so it goes to the same column of the row that stays (`survivingCell` in move-rescue.ts).
      expect(tablesConverged(editors)).toEqual(['CB|Z']);
    });

  for (const order of THREE_ORDERS)
    test(`typing in the merged cell survives when both rows go (client order ${names(order)})`, async () => {
      const { editors, sync } = await peers(threeRows(), order);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a);
      caret(a, 0, 0);
      expect(a.input('text', 'x').ok).toBe(true);
      deleteRow(b, 1);
      deleteRow(c, 0);
      sync.flush();
      expect(tablesConverged(editors)).toEqual(['Cx|Z']);
    });

  for (const order of THREE_ORDERS)
    test(`undo of each delete brings its row back, and redo removes it again (client order ${names(order)})`, async () => {
      const { editors, sync } = await peers(threeRows(), order);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a);
      deleteRow(b, 1);
      deleteRow(c, 0);
      sync.flush();
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      // C restores its row and "A". B's independent deletion of "B" still applies.
      for (const e of editors) expect(grid(e)).toEqual(['A|X', 'C|Z']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      // The lower row comes back as the covered cell of the merge, with its empty paragraph.
      expect(tablesConverged(editors)).toEqual(['AB|X', '|Y', 'C|Z']);
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      expect(tablesConverged(editors)).toEqual(['C|Z']);
    });
});
