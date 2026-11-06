/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A vertical merge and a concurrent delete of one of its covered rows (AGENTS.md,
// "Collaboration verification"): two editors on linked Y.Docs, the link paused, both client
// orders. The merge's span then reaches the row below the deleted one. Each replica saves with
// no further local write: both save the same bytes, the saved deck passes the part checks and
// reopens, and the row below shows again. The stores keep the room's state.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  textOf,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import { META_KEY, SCHEMA_VERSION } from '../collaboration/yjs/schema.ts';
import { isStructureEdited } from '../../../core/src/store/package/table-marks.ts';
import { IN, TWO_ORDERS, all, kids, peers, root, slide1 } from './lineage-support.ts';

/**
 * A 3 x 3 table, one letter per cell; `blank` empties the cells it names ("row,column"), and
 * `attrs` adds attributes to the cells it names.
 */
const deck = (
  blank: readonly string[] = [],
  attrs: Readonly<Record<string, string>> = {}
): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: 3 * 370_840,
          columns: [IN, IN, IN],
          tblPr: '<a:tblPr firstRow="1" bandRow="1"/>',
          rows: ['ABC', 'DEF', 'GHI']
            .map(
              (r, i) =>
                `<a:tr h="370840">${[...r]
                  .map((c, j) =>
                    cell(blank.includes(`${i},${j}`) ? '' : c, { attrs: attrs[`${i},${j}`] ?? '' })
                  )
                  .join('')}</a:tr>`
            )
            .join(''),
        }),
      },
    ],
  });

/** The slide's graphic frame (PresentationML, so `all` does not find it). */
function frameOf(e: PptxEditorInstance): Element {
  const find = (el: Element): Element | undefined => {
    for (const c of kids(el)) {
      const hit = c.kind === 'graphicFrame' ? c : find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  return find(root(e))!;
}
const rowsOf = (tbl: Element): Element[][] =>
  all(tbl, 'tr').map((tr) => kids(tr).filter((c) => c.local === 'tc'));
const merge = (tc: Element): string =>
  tc.attrs
    .filter((a) => ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].includes(a.local))
    .map((a) => `${a.local}=${a.value}`)
    .join(',');

/** Merge rows `r0`..`r1` of the first column on `e`. */
function mergeFirstColumn(e: PptxEditorInstance, r0: number, r1: number): void {
  e.setSelection({
    kind: 'cells',
    nodeId: frameOf(e).id,
    anchor: { row: r0, column: 0 },
    focus: { row: r1, column: 0 },
  });
  expect(e.exec({ type: 'mergeCells' }).ok).toBe(true);
}

/** Delete row `row` on `e`, from a caret in its last column. */
function deleteRow(e: PptxEditorInstance, row: number): void {
  e.setSelection({
    kind: 'text',
    nodeId: frameOf(e).id,
    cell: { row, column: 2 },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
  expect(e.exec({ type: 'deleteRow' }).ok).toBe(true);
}

/**
 * Each replica saves with no local write before it. The saves are equal, pass the part checks,
 * and reopen; the stores do not change. Returns the saved table: per row, each cell's text and
 * merge attributes.
 */
function savedEqual(a: PptxEditorInstance, b: PptxEditorInstance, valid = true): string[] {
  const before = canonicalPackageFingerprint(a.session!.store.package);
  expect(canonicalPackageFingerprint(b.session!.store.package)).toBe(before);
  const revisions = [a.session!.store.revision, b.session!.store.revision];
  const [bytesA, bytesB] = [a.save()!, b.save()!];
  expect([a.session!.store.revision, b.session!.store.revision]).toEqual(revisions);
  expect(canonicalPackageFingerprint(a.session!.store.package)).toBe(before);
  const [ra, rb] = [readPptx(bytesA), readPptx(bytesB)];
  if (!ra.ok || !rb.ok) throw new Error('reopen failed');
  expect(canonicalPackageFingerprint(ra.pkg)).toBe(canonicalPackageFingerprint(rb.pkg));
  expect(semanticDigest(ra.pkg)).toBe(semanticDigest(rb.pkg));
  if (valid) expect(packageViolations(ra.pkg)).toEqual([]);
  expect(new TextDecoder().decode(bytesA)).not.toContain('table-structure-edit');
  const tbl = all(xmlPartOf(ra.pkg, slide1(a))!.root, 'tbl')[0]!;
  return rowsOf(tbl).map((row) =>
    row.map((tc) => (merge(tc) ? `${textOf(tc)}[${merge(tc)}]` : textOf(tc))).join('|')
  );
}

const orderName = (order: readonly number[]) => (order[0]! < order[1]! ? 'A, B' : 'B, A');

describe('a vertical merge and a concurrent delete of one of its covered rows', () => {
  for (const order of TWO_ORDERS)
    test(`the merged cell ends at its last row and the next row shows (client order ${orderName(order)})`, async () => {
      const { editors, sync } = await peers(deck(), order);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a, 0, 1);
      deleteRow(b, 1);
      sync.flush();
      // The room keeps the merge's span, so the store still covers "G".
      expect(merge(rowsOf(all(root(a), 'tbl')[0]!)[0]![0]!)).toBe('rowSpan=2');
      expect(savedEqual(a, b)).toEqual(['AD|B|C', 'G|H|I']);
    });

  for (const order of TWO_ORDERS)
    test(`a merge of three rows that loses its middle row spans the two left (client order ${orderName(order)})`, async () => {
      const { editors, sync } = await peers(deck(), order);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a, 0, 2);
      deleteRow(b, 1);
      sync.flush();
      expect(savedEqual(a, b)).toEqual(['ADG[rowSpan=2]|B|C', '[vMerge=1]|H|I']);
    });

  for (const order of TWO_ORDERS)
    test(`a blank cell below stays covered, with the marker the next table command writes (client order ${orderName(order)})`, async () => {
      const { editors, sync } = await peers(deck(['2,0']), order);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a, 0, 1);
      deleteRow(b, 1);
      sync.flush();
      const saved = savedEqual(a, b);
      expect(saved).toEqual(['AD[rowSpan=2]|B|C', '[vMerge=1]|H|I']);
      // The next table command writes the same merge markup into the room.
      a.setSelection({
        kind: 'text',
        nodeId: frameOf(a).id,
        cell: { row: 1, column: 1 },
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 1 },
      });
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(rowsOf(all(root(b), 'tbl')[0]!).map((row) => row.map(merge).join('|'))).toEqual([
        'rowSpan=2||',
        'vMerge=1||',
      ]);
      expect(savedEqual(a, b)).toEqual(saved);
    });
});

describe('the save repair across replicas', () => {
  for (const order of TWO_ORDERS)
    test(`a merge and a concurrent delete of the column of its covered cell save the next cell visible (client order ${orderName(order)})`, async () => {
      const { editors, sync } = await peers(deck(), order);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      a.setSelection({
        kind: 'cells',
        nodeId: frameOf(a).id,
        anchor: { row: 0, column: 0 },
        focus: { row: 0, column: 1 },
      });
      expect(a.exec({ type: 'mergeCells' }).ok).toBe(true);
      b.setSelection({
        kind: 'text',
        nodeId: frameOf(b).id,
        cell: { row: 2, column: 1 },
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 0 },
      });
      expect(b.exec({ type: 'deleteColumn' }).ok).toBe(true);
      sync.flush();
      expect(savedEqual(a, b)).toEqual(['AB|C', 'D|F', 'G|I']);
    });

  test('an undo of the merge on one replica saves the table as the file wrote it on both', async () => {
    // The file's last row holds an orphan marker: a save keeps it while no structure mark is on.
    const { editors, sync } = await peers(deck([], { '2,2': ' vMerge="1"' }), TWO_ORDERS[0]!);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    const file = savedEqual(a, b, false);
    mergeFirstColumn(a, 0, 1);
    sync.flush();
    expect(savedEqual(a, b)).toEqual(['AD[rowSpan=2]|B|C', '[vMerge=1]|E|F', 'G|H|I']);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    sync.flush();
    expect(savedEqual(a, b, false)).toEqual(file);
  });

  test('a replica that opened another deck saves the table as the room holds it', async () => {
    const docs = [new Y.Doc(), new Y.Doc()] as const;
    docs[0].clientID = 1_001;
    docs[1].clientID = 2_002;
    const link = (from: Y.Doc, to: Y.Doc) =>
      from.on('update', (update: Uint8Array, origin: unknown) => {
        if (origin !== 'link') Y.applyUpdate(to, update, 'link');
      });
    link(docs[0], docs[1]);
    link(docs[1], docs[0]);
    const open = async (doc: Y.Doc, bytes: Uint8Array) => {
      const session = createYjsCollaborationSession({
        doc,
        documentId: 'deck',
        identity: { actorId: String(doc.clientID), name: String(doc.clientID) },
      });
      const editor = createPptxEditor({
        fixedMeasurer: true,
        substitutes: false,
        modules: [collaborationModule({ session })],
      });
      expect((await editor.load(bytes)).ok).toBe(true);
      return editor;
    };
    // The room starts from a deck with an orphan marker; the second replica opens another deck.
    const a = await open(docs[0], deck([], { '2,2': ' vMerge="1"' }));
    const b = await open(docs[1], deck());
    expect(savedEqual(a, b, false)).toEqual(['A|B|C', 'D|E|F', 'G|H|I[vMerge=1]']);
  });
});

/**
 * The table as `e` draws it: per grid row, each drawn cell as its text with its spans when
 * they are not 1 x 1, and `·` for a covered position.
 */
function drawn(e: PptxEditorInstance): string[] {
  const item = e.session!.layout(0)!.items.find((i) => i.element.kind === 'graphicFrame')!;
  const table = item.element.kind === 'graphicFrame' ? item.element.table! : undefined;
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out: string[] = [];
  for (const c of table!.cells) {
    if (c.col === 0) out.push('');
    const text = textOf(byId.get(c.nodeId)!);
    const spans = c.gridSpan > 1 || c.rowSpan > 1 ? `[${c.rowSpan}x${c.gridSpan}]` : '';
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : text + spans}`;
  }
  return out;
}

/** The table that a fresh editor draws for `e`'s save. */
async function drawnAfterSave(e: PptxEditorInstance): Promise<string[]> {
  const reopened = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await reopened.load(e.save()!)).ok).toBe(true);
  return drawn(reopened);
}

describe('the canvas draws what the save writes', () => {
  for (const order of TWO_ORDERS)
    test(`after a merge and a concurrent row delete, both replicas draw the next row (client order ${orderName(order)})`, async () => {
      const { editors, sync } = await peers(deck(), order);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a, 0, 1);
      deleteRow(b, 1);
      sync.flush();
      const revisions = [a.session!.store.revision, b.session!.store.revision];
      expect(drawn(a)).toEqual(['AD|B|C', 'G|H|I']);
      expect(drawn(b)).toEqual(drawn(a));
      expect(await drawnAfterSave(a)).toEqual(drawn(a));
      expect(await drawnAfterSave(b)).toEqual(drawn(b));
      // Drawing wrote nothing: the store still holds the room's span.
      expect([a.session!.store.revision, b.session!.store.revision]).toEqual(revisions);
      expect(merge(rowsOf(all(root(a), 'tbl')[0]!)[0]![0]!)).toBe('rowSpan=2');
    });

  for (const order of TWO_ORDERS)
    test(`a blank cell below the merge stays covered on the canvas, as it saves (client order ${orderName(order)})`, async () => {
      const { editors, sync } = await peers(deck(['2,0']), order);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      mergeFirstColumn(a, 0, 1);
      deleteRow(b, 1);
      sync.flush();
      expect(drawn(a)).toEqual(['AD[2x1]|B|C', '·|H|I']);
      expect(drawn(b)).toEqual(drawn(a));
      expect(await drawnAfterSave(a)).toEqual(drawn(a));
    });

  test('typing into the row that shows again writes into that row, on both replicas', async () => {
    const { editors, sync } = await peers(deck(), TWO_ORDERS[0]!);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    sync.pause();
    mergeFirstColumn(a, 0, 1);
    deleteRow(b, 1);
    sync.flush();
    b.setSelection({
      kind: 'text',
      nodeId: frameOf(b).id,
      cell: { row: 1, column: 0 },
      anchor: { paragraph: 0, offset: 1 },
      focus: { paragraph: 0, offset: 1 },
    });
    expect(b.input('text', 'x').ok).toBe(true);
    sync.flush();
    expect(drawn(a)).toEqual(['AD|B|C', 'Gx|H|I']);
    expect(drawn(b)).toEqual(drawn(a));
    expect(savedEqual(a, b)).toEqual(['AD|B|C', 'Gx|H|I']);
  });

  test('a copy of the table pastes as the canvas draws it, without a mark', async () => {
    const { editors, sync } = await peers(deck(), TWO_ORDERS[0]!);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    sync.pause();
    mergeFirstColumn(a, 0, 1);
    deleteRow(b, 1);
    sync.flush();
    a.setSelection({ kind: 'shapes', nodeIds: [frameOf(a).id] });
    const payload = a.copy();
    expect(JSON.stringify(payload)).not.toContain('table-structure-edit');
    expect(a.exec({ type: 'paste', payload: payload! }).ok).toBe(true);
    sync.flush();
    const tables = all(root(b), 'tbl').map((tbl) =>
      rowsOf(tbl).map((row) => row.map((tc) => textOf(tc) + (merge(tc) ? `[${merge(tc)}]` : '')))
    );
    expect(tables[1]).toEqual([
      ['AD', 'B', 'C'],
      ['G', 'H', 'I'],
    ]);
    expect(savedEqual(a, b)).toEqual(['AD|B|C', 'G|H|I']);
  });
});

describe('the room format names structure marks', () => {
  test('a structure command writes marks only into a room of the current format', async () => {
    const { editors, sync } = await peers(deck(), TWO_ORDERS[0]!);
    const [a] = editors as [PptxEditorInstance, PptxEditorInstance];
    mergeFirstColumn(a, 0, 1);
    sync.flush();
    expect(SCHEMA_VERSION).toBe(5);
    expect(isStructureEdited(frameOf(a))).toBe(true);
  });

  for (const [name, schema] of [
    ['an older format (2), whose peers would save marks', 2],
    ['an older format (3), whose peers would hide merged text', 3],
    ['format 4, whose peers repair no undo by its listings and marks', 4],
    ['a newer format', SCHEMA_VERSION + 1],
  ] as const)
    test(`a replica refuses a room of ${name}, keeps its own deck, and saves no mark`, async () => {
      // A room that a peer of the current format filled, with a marked table.
      const { editors, docs } = await peers(deck(), [1_001]);
      mergeFirstColumn(editors[0]!, 0, 1);
      const doc = docs[0]!;
      expect(doc.getMap(META_KEY).get('schema')).toBe(SCHEMA_VERSION);
      doc.getMap(META_KEY).set('schema', schema);
      const session = createYjsCollaborationSession({
        doc,
        documentId: 'deck',
        identity: { actorId: 'late', name: 'late' },
      });
      const late = createPptxEditor({
        fixedMeasurer: true,
        substitutes: false,
        modules: [collaborationModule({ session })],
      });
      expect((await late.load(deck(['0,1']))).ok).toBe(true);
      expect(session.statusSnapshot()).toMatchObject({
        status: 'error',
        reason: { code: 'schema-version-mismatch' },
      });
      expect(session.attached).toBe(false);
      const saved = late.save()!;
      expect(new TextDecoder().decode(saved)).not.toContain('table-structure-edit');
      const r = readPptx(saved);
      if (!r.ok) throw new Error('reopen failed');
      // Its own deck: the cell "B" it opened blank, and no merge from the room.
      const tbl = all(xmlPartOf(r.pkg, slide1(late))!.root, 'tbl')[0]!;
      expect(rowsOf(tbl)[0]!.map((tc) => textOf(tc) + merge(tc))).toEqual(['A', '', 'C']);
    });
});

describe('structure marks and undo in a room', () => {
  /** `first` inserts a row below row 0, `second` merges rows 0 to 2 over it, `first` undoes. */
  async function undoUnderMerge(swap: boolean) {
    const { editors, sync } = await peers(deck(), TWO_ORDERS[0]!);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    const [first, second] = swap ? [b, a] : [a, b];
    first.setSelection({
      kind: 'text',
      nodeId: frameOf(first).id,
      cell: { row: 0, column: 2 },
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    expect(first.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    sync.flush();
    mergeFirstColumn(second, 0, 2);
    sync.flush();
    // The undo takes the inserted row and its own mark; the merge keeps its mark, so a save
    // ends the merge's span before "G".
    expect(first.exec({ type: 'undo' }).ok).toBe(true);
    sync.flush();
    const saved = savedEqual(a, b);
    expect(saved).toHaveLength(3);
    expect(saved[2]).toBe('G|H|I');
  }

  test('an undo of a row insert that a peer merged across saves valid merge markers', () =>
    undoUnderMerge(false));

  test('the same with the peers swapped', () => undoUnderMerge(true));
});
