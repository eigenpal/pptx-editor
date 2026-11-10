/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A cell merge while peers change the same table's structure or type in it (AGENTS.md,
// "Collaboration verification"): editors on linked Y.Docs, the links paused, each client
// order, undo and redo. The histories are the shrunk failures of the seeded table fuzzer
// (eigenpal/pptx-editor#1, #2). Merged text shows in the drawn cell that covers the
// cell it sits in (`merged-cells.ts`), and a merge leaves no hidden copy of it. Every case
// checks that each character shows once on the canvas, that no character a peer did not
// remove is lost, that the replicas hold the same package, and that the saved deck passes the
// part checks and reopens to the same deck on every replica.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  textOf,
} from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import { IN, all, kids, peers, root, sessionsOf } from './lineage-support.ts';

/**
 * A 3 x 3 table, one letter per cell; `blank` empties the cells it names ("row,column"), and
 * `bodiless` gives the cells it names no text body.
 */
const deck = (blank: readonly string[] = [], bodiless: readonly string[] = []): Uint8Array =>
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
                    bodiless.includes(`${i},${j}`)
                      ? '<a:tc><a:tcPr/></a:tc>'
                      : cell(blank.includes(`${i},${j}`) ? '' : c)
                  )
                  .join('')}</a:tr>`
            )
            .join(''),
        }),
      },
    ],
  });

const TWO = [
  [1_001, 2_002],
  [2_002, 1_001],
] as const;
const THREE = [
  [1_001, 2_002, 3_003],
  [3_003, 2_002, 1_001],
  [2_002, 3_003, 1_001],
] as const;
const named = (ids: readonly number[]) => ids.map((id) => id / 1_001).join(', ');

/** The slide's table frame. */
function frameId(e: PptxEditorInstance): string {
  const find = (el: ReturnType<typeof root>): string | undefined => {
    for (const c of kids(el)) {
      const hit = c.kind === 'graphicFrame' ? c.id : find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  return find(root(e))!;
}

/** The table that layout draws. */
function layoutTable(e: PptxEditorInstance) {
  const item = e.session!.layout(0)?.items.find((i) => i.element.kind === 'graphicFrame');
  return item?.element.kind === 'graphicFrame' ? item.element.table : undefined;
}

/** Per grid row, each drawn cell's text with its spans when not 1 x 1, `·` when covered. */
function drawn(e: PptxEditorInstance): string[] {
  const t = layoutTable(e);
  if (!t) return [];
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out: string[] = [];
  for (const c of t.cells) {
    if (c.col === 0) out.push('');
    const tc = byId.get(c.nodeId);
    const spans = c.gridSpan > 1 || c.rowSpan > 1 ? `[${c.rowSpan}x${c.gridSpan}]` : '';
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : (tc ? textOf(tc) : '?') + spans}`;
  }
  return out;
}

/** The text of the cells the canvas draws, in reading order. */
function shown(e: PptxEditorInstance): string {
  const ids = new Set((layoutTable(e)?.cells ?? []).filter((c) => !c.covered).map((c) => c.nodeId));
  return all(root(e), 'tc')
    .filter((tc) => ids.has(tc.id))
    .map((tc) => textOf(tc))
    .join('');
}

function at(e: PptxEditorInstance, row: number, column: number, offset = 0): void {
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset },
    focus: { paragraph: 0, offset },
  });
}

function merge(e: PptxEditorInstance, r0: number, c0: number, r1: number, c1: number): void {
  e.setSelection({
    kind: 'cells',
    nodeId: frameId(e),
    anchor: { row: r0, column: c0 },
    focus: { row: r1, column: c1 },
  });
  expect(e.exec({ type: 'mergeCells' }).ok).toBe(true);
}

function run(e: PptxEditorInstance, row: number, column: number, command: object): void {
  at(e, row, column);
  expect(e.exec(command as Parameters<PptxEditorInstance['exec']>[0]).ok).toBe(true);
}

/**
 * The replicas hold the same package, every session applied every change, the canvas shows
 * `text` (each character once), and each replica's save passes the part checks and reopens to
 * the same deck, which a second save keeps. Returns the drawn table.
 */
function settled(editors: readonly PptxEditorInstance[], text: string): string[] {
  for (const s of sessionsOf(editors)) expect(s.statusSnapshot().status).toBe('ready');
  const [first, ...rest] = editors;
  const fingerprint = (e: PptxEditorInstance) =>
    canonicalPackageFingerprint(e.session!.store.package);
  for (const e of rest) expect(fingerprint(e)).toBe(fingerprint(first!));
  for (const e of editors) expect([...shown(e)].sort().join('')).toBe([...text].sort().join(''));
  const saves = editors.map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('the saved deck does not reopen');
    expect(packageViolations(r.pkg)).toEqual([]);
    expect(semanticDigest(r.pkg)).toBe(canonicalPackageFingerprint(r.pkg));
    return canonicalPackageFingerprint(r.pkg);
  });
  for (const s of saves) expect(s).toBe(saves[0]!);
  return drawn(first!);
}

describe('a merge and a concurrent delete of its first cell keep the merged text (#1)', () => {
  for (const ids of TWO)
    test(`a merge of two columns and a delete of the first column (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(a, 0, 0, 2, 1);
      run(b, 0, 0, { type: 'deleteColumn' });
      sync.flush();
      // "A", "D", and "G" went with the column, as B asked. "B", "E", and "H" show once.
      expect(settled(editors, 'BEHCFI')).toEqual(['BEH[3x1]|C', '·|F', '·|I']);
      // The undo of the delete brings the column back; the merged cell holds all six.
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABDEGHCFI');
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      expect(settled(editors, 'BEHCFI')).toEqual(['BEH[3x1]|C', '·|F', '·|I']);
      // The undo of the merge gives every letter its own cell again.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(settled(editors, 'BEHCFI')).toEqual(['B|C', 'E|F', 'H|I']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'BEHCFI');
    });

  for (const ids of THREE)
    test(`a merge of two rows and a delete of the first row, three peers (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      run(c, 0, 2, { type: 'deleteRow' });
      merge(b, 0, 2, 2, 1);
      sync.flush();
      // "B" and "C" went with the row; "E", "F", "H", and "I" show once. Each cell of the
      // first column of the merge spans the merge's width, so each row draws one wide cell.
      expect(settled(editors, 'DEFGHI')).toEqual(['D|EF[1x2]|·', 'G|HI[1x2]|·']);
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'DEFGHI');
    });

  for (const ids of THREE)
    test(`a merge, a delete of the merged cell's row, and a delete of a column the merged text left (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // A merges "D", "E", "G", and "H", then deletes the first column: the merged cell keeps
      // its text. B deletes the row of "D", "E", and "F" without seeing the merge.
      merge(a, 1, 0, 2, 1);
      run(b, 1, 0, { type: 'deleteRow' });
      run(a, 0, 0, { type: 'deleteColumn' });
      sync.flush();
      // "G" left its cell before A deleted that cell, so it goes with "H".
      expect(settled(editors, 'BCGHI')).toEqual(['B|C', 'GH|I']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'BCDEGHFI');
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      expect(settled(editors, 'BCGHI')).toEqual(['B|C', 'GH|I']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCGHI');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'BCGHI');
    });

  for (const ids of THREE)
    test(`a merge, a delete of the merged cell's row, and a concurrent delete of a column of the merge (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // B deletes the row of "D", and C deletes the column of "A", "D", and "G": neither sees
      // A's merge, so "G" goes with C's column.
      merge(a, 1, 0, 2, 1);
      run(b, 1, 0, { type: 'deleteRow' });
      run(c, 0, 0, { type: 'deleteColumn' });
      sync.flush();
      expect(settled(editors, 'BCHI')).toEqual(['B|C', 'H|I']);
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCGHI');
    });

  for (const ids of THREE)
    test(`a merge into a new row and a concurrent undo of the row insert (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // A inserts a row above the first row, and B sees it. A undoes the insert while B merges
      // the first column, from the new row down: the merge moves "A", "D", and "G" into the
      // new row's cell, which A's undo takes out of the table.
      run(a, 0, 0, { type: 'insertRow', where: 'above' });
      sync.deliverTo(docs[1]!);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      merge(b, 0, 0, 3, 0);
      sync.flush();
      // The undo did not see the merge, so the merged text goes back to its cells.
      settled(editors, 'ABCDEFGHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of THREE)
    test(`a merge whose first cell a peer deletes, over a cell with no text body (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck([], ['1,1']), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // A merges "D", the cell with no text body, "G", and "H". B deletes the first column
      // without seeing the merge, so the cell with no text body becomes the merged cell.
      merge(a, 1, 0, 2, 1);
      run(b, 0, 0, { type: 'deleteColumn' });
      sync.flush();
      // That cell takes the text body that holds "H", so "H" shows.
      expect(settled(editors, 'BCFHI')).toEqual(['B|C', 'H[2x1]|F', '·|I']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDFGHI');
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'BCFHI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(settled(editors, 'BCFHI')).toEqual(['B|C', '|F', 'H|I']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'BCFHI');
    });

  for (const ids of TWO)
    test(`a staged merge over a cell with no text body and a delete of its first column (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck([], ['1,1']), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(a, 1, 0, 2, 1);
      run(b, 0, 0, { type: 'deleteColumn' });
      // B gets the merge first, and types into the merged cell before A gets the delete.
      sync.deliverTo(docs[1]!);
      at(b, 1, 0);
      expect(b.input('text', 'x').ok).toBe(true);
      sync.flush();
      settled(editors, 'BCFHIx');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'BCFHI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'BCFHI');
    });

  for (const ids of THREE)
    test(`two overlapping merges and a split show every letter once (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(c, 1, 1, 1, 2);
      merge(b, 1, 2, 2, 2);
      run(b, 1, 2, { type: 'splitCell', rows: 1, cols: 2 });
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });
});

describe('a merge and a concurrent structure change inside it show merged text once', () => {
  for (const ids of THREE)
    test(`a merge of a column and a row insert inside it (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, , c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(a, 2, 1, 0, 1);
      run(c, 2, 0, { type: 'insertRow', where: 'above' });
      sync.flush();
      expect(settled(editors, 'ABEHCDFGI')).toEqual(['A|BEH[3x1]|C', 'D|·|F', '|·|', 'G||I']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABEHCDFGI');
    });

  for (const ids of TWO)
    test(`a merge and a split of a cell inside it (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(b, 0, 1, 1, 2);
      run(a, 0, 2, { type: 'splitCell', rows: 2, cols: 2 });
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of TWO)
    test(`a merge of a row and a column insert inside it (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(a, 0, 0, 0, 2);
      run(b, 0, 1, { type: 'insertColumn', where: 'left' });
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of TWO)
    test(`two merges whose spans together cover cells that neither merged (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // A merges the first two rows, and B merges the first column. Together their spans
      // can cover "H" and "I", which neither moved (the client order picks the spans): the
      // merged cell then shows them.
      merge(a, 0, 0, 1, 2);
      merge(b, 0, 0, 2, 0);
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      // A later edit keeps the text in the merged cell, also after an undo uncovers the cells.
      at(a, 0, 0, 1);
      expect(a.input('text', 'x').ok).toBe(true);
      sync.flush();
      settled(editors, 'AxBCDEFGHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'AxBCDEFGHI');
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'AxBCDEFGHI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of THREE)
    test(`two staged merges whose spans together cover a cell (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // C sees A's merge, then B's, before A and B see each other's.
      merge(a, 0, 0, 1, 2);
      sync.deliverTo(docs[2]!);
      merge(b, 0, 0, 2, 0);
      sync.deliverTo(docs[2]!);
      expect([...shown(c)].sort().join('')).toBe('ABCDEFGHI');
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of TWO)
    test(`two overlapping merges (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      merge(a, 2, 2, 0, 1);
      merge(b, 0, 2, 1, 1);
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });
});

describe('text typed into a cell that a peer merges at the same time is kept (#2)', () => {
  for (const ids of THREE)
    test(`typing into the empty first cell of a merge (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      run(a, 0, 2, { type: 'insertRow', where: 'above' });
      sync.deliverTo(docs[1]!);
      merge(b, 1, 1, 0, 0);
      at(a, 0, 0);
      expect(a.input('text', 'x').ok).toBe(true);
      sync.flush();
      // The merged cell shows "x" with "A" and "B".
      expect(settled(editors, 'xABCDEFGHI')[0]).toBe('xAB[2x2]|·|');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(settled(editors, 'ABCDEFGHI')[0]).toBe('AB[2x2]|·|');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'xABCDEFGHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'xABCDEFGHI');
    });

  for (const ids of THREE)
    test(`typing into an empty cell that a peer's merge covers (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(['0,2']), ids);
      const [, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      at(b, 0, 2);
      expect(b.input('text', 'x').ok).toBe(true);
      merge(c, 0, 0, 0, 2);
      sync.flush();
      // The merged cell shows "x" after "A" and "B".
      expect(settled(editors, 'ABxDEFGHI')[0]).toBe('ABx[1x3]|·|·');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(settled(editors, 'ABDEFGHI')[0]).toBe('AB[1x3]|·|·');
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABxDEFGHI');
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      expect(settled(editors, 'ABxDEFGHI')[0]).toBe('A|B|x');
    });

  for (const ids of THREE)
    test(`typing into a cell that a split uncovered and a peer merges again (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      // A merges the first two rows, and B splits the merged cell into its two rows again. The
      // second row's cell is empty: A's merge saw its empty paragraph hidden there.
      merge(a, 0, 0, 1, 2);
      sync.flush();
      run(b, 0, 0, { type: 'splitCell', rows: 2, cols: 1 });
      sync.flush();
      sync.pause();
      // A merges the two rows again while C types into the second row's cell.
      merge(a, 1, 1, 0, 2);
      at(c, 1, 0);
      expect(c.input('text', 'x').ok).toBe(true);
      sync.flush();
      // The merged cell shows "x" with the other text.
      settled(editors, 'ABCDEFGHIx');
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHIx');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHIx');
    });

  for (const ids of THREE)
    test(`a merge into an empty cell that a peer's merge covers (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, , c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      run(a, 1, 2, { type: 'insertRow', where: 'above' });
      sync.deliverTo(docs[2]!);
      // C merges "B" with the new empty cell below it; A merges the new empty cell with the
      // cells of "E" and "F", so "E" and "F" move into a cell that C covers.
      merge(c, 1, 1, 0, 1);
      merge(a, 2, 1, 1, 2);
      sync.flush();
      settled(editors, 'ABCDEFGHI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of TWO)
    test(`typing into an empty paragraph that a peer removes (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      // An empty second paragraph in the first cell.
      at(a, 0, 0, 1);
      expect(a.input('paragraph').ok).toBe(true);
      sync.flush();
      sync.pause();
      a.setSelection({
        kind: 'text',
        nodeId: frameId(a),
        cell: { row: 0, column: 0 },
        anchor: { paragraph: 1, offset: 0 },
        focus: { paragraph: 1, offset: 0 },
      });
      expect(a.input('deleteBackward').ok).toBe(true);
      b.setSelection({
        kind: 'text',
        nodeId: frameId(b),
        cell: { row: 0, column: 0 },
        anchor: { paragraph: 1, offset: 0 },
        focus: { paragraph: 1, offset: 0 },
      });
      expect(b.input('text', 'y').ok).toBe(true);
      sync.flush();
      expect(settled(editors, 'AyBCDEFGHI')[0]).toBe('Ay|B|C');
      expect(all(all(root(a), 'tc')[0]!, 'p').map((p) => textOf(p))).toEqual(['A', 'y']);
    });
});

describe('an undo of an insert keeps the text a peer typed into a new cell (#109)', () => {
  for (const ids of THREE)
    test(`a column insert (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      run(a, 0, 2, { type: 'insertColumn', where: 'right' });
      sync.flush();
      sync.pause();
      // B types into the new cell of the last row while A undoes the insert.
      at(b, 2, 3);
      expect(b.input('text', 'x').ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      // The column goes, and "x" shows in the cell it followed.
      expect(settled(editors, 'ABCDEFGHIx')).toEqual(['A|B|C', 'D|E|F', 'G|H|Ix']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(settled(editors, 'ABCDEFGHIx')).toEqual(['A|B|C|', 'D|E|F|', 'G|H|I|x']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHIx');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(settled(editors, 'ABCDEFGHI')).toEqual(['A|B|C', 'D|E|F', 'G|H|I']);
    });

  for (const ids of THREE)
    test(`a row insert (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      run(a, 2, 0, { type: 'insertRow', where: 'below' });
      sync.flush();
      sync.pause();
      at(b, 3, 1);
      expect(b.input('text', 'x').ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      // The row goes, and "x" shows in the same column of the row it followed.
      expect(settled(editors, 'ABCDEFGHIx')).toEqual(['A|B|C', 'D|E|F', 'G|Hx|I']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(settled(editors, 'ABCDEFGHIx')).toEqual(['A|B|C', 'D|E|F', 'G|H|I', '|x|']);
    });

  for (const ids of TWO)
    test(`a split, with the undo delivered in stages (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      run(a, 1, 1, { type: 'splitCell', rows: 1, cols: 2 });
      sync.flush();
      sync.pause();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      at(b, 1, 2);
      expect(b.input('text', 'x').ok).toBe(true);
      // A gets B's text first, then B gets the undo.
      sync.deliverTo(docs[0]!);
      sync.flush();
      expect(settled(editors, 'ABCDEFGHIx')).toEqual(['A|B|C', 'D|Ex|F', 'G|H|I']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });

  for (const ids of TWO)
    test(`a merge of the cell that shows the text (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      run(a, 0, 0, { type: 'insertColumn', where: 'left' });
      sync.flush();
      sync.pause();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      at(b, 1, 0);
      expect(b.input('text', 'x').ok).toBe(true);
      // B gets the undo, sees "x" in "D", and merges the first column: "x" goes with "D".
      sync.deliverTo(docs[1]!);
      merge(b, 0, 0, 2, 0);
      sync.flush();
      expect(settled(editors, 'ABCDEFGHIx')).toEqual(['ADxG[3x1]|B|C', '·|E|F', '·|H|I']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHIx');
    });
});

describe('merged text whose old and new cell both go shows in a cell that stays (#1)', () => {
  for (const ids of THREE)
    test(`a merge, a delete of the row it left, and a delete of the merged cell's row (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // B merges "C" and "F", then deletes the row of "F". A deletes the row of "C" without
      // seeing the merge: "F" left its row before B deleted it, and A did not see it in "C".
      merge(b, 0, 2, 1, 2);
      run(b, 1, 1, { type: 'deleteRow' });
      run(a, 0, 2, { type: 'deleteRow' });
      sync.flush();
      expect(settled(editors, 'FGHI')).toEqual(['G|H|IF']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCFGHI');
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'FGHI');
    });

  for (const ids of THREE)
    test(`two merges move one paragraph into a covered cell over a cell with no text body (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck([], ['1,1']), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      run(a, 0, 1, { type: 'deleteRow' });
      run(a, 1, 1, { type: 'insertRow', where: 'above' });
      sync.flush();
      sync.pause();
      // B deletes the first column and merges the new row's cell with "H". C merges the first
      // two columns of all rows. Both move "H"; the cell with no text body becomes C's merged
      // cell, and the cell that holds "H" for B is covered by it.
      run(b, 2, 0, { type: 'deleteColumn' });
      merge(c, 2, 1, 0, 0);
      merge(b, 1, 0, 2, 0);
      sync.flush();
      settled(editors, 'FHI');
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'FHI');
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'FHI');
    });
});

describe('a text body whose old and new cell both go shows in a cell that stays (#1)', () => {
  for (const ids of THREE)
    test(`a merge of a row, a delete of its first column, and a delete of the second (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, , c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // A merges "D", "E", and "F", then deletes the first column: the merged text moves to the
      // cell of "E". C deletes the column of "E" without seeing the merge.
      merge(a, 1, 0, 1, 2);
      run(a, 0, 0, { type: 'deleteColumn' });
      run(c, 1, 1, { type: 'deleteColumn' });
      sync.flush();
      // C deleted "E", but never saw "D" or "F" in that cell. Only those unseen characters
      // move to the cell that stays. C's undo restores "E" with the rest of its column.
      settled(editors, 'CDFI');
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'BCDEFHI');
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'CDFI');
    });
});

describe('an undo of a merge, then a delete of a cell it gave text back to', () => {
  for (const ids of THREE)
    test(`replicas agree and the text goes with the cell (client order ${named(ids)})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, , c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      // A inserts a column that C does not see yet. C merges the first two columns and undoes
      // the merge, then deletes the column of "B", "E", and "H", where the undo put them back.
      run(a, 0, 1, { type: 'insertColumn', where: 'left' });
      merge(c, 2, 1, 0, 0);
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(drawn(c)).toEqual(['A||B|C', 'D||E|F', 'G||H|I']);
      run(c, 1, 2, { type: 'deleteColumn' });
      sync.deliverTo(docs[0]!);
      sync.flush();
      expect(settled(editors, 'ACDFGI')).toEqual(['A||C', 'D||F', 'G||I']);
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDEFGHI');
    });
});

describe('an undo of a merge keeps the text when a peer removes the cell it goes back to', () => {
  for (const ids of THREE)
    test(`a concurrent delete of the column of the merged text (client order ${named(ids)})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      merge(a, 0, 0, 0, 1);
      sync.flush();
      sync.pause();
      // A undoes the merge, so "B" goes back to its cell. B deletes that column at the same
      // time: B saw "B" in the merged cell, not in the column it deletes.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      run(b, 1, 1, { type: 'deleteColumn' });
      sync.flush();
      expect(settled(editors, 'ABCDFGI')).toEqual(['AB|C', 'D|F', 'G|I']);
      // A's redo merges again: "B" stays in the merged cell.
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      settled(editors, 'ABCDFGI');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      settled(editors, 'ABCDFGI');
    });
});
