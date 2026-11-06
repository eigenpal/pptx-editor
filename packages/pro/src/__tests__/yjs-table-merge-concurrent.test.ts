/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
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

/** A 3 x 3 table, one letter per cell; `blank` empties the cells it names ("row,column"). */
const deck = (blank: readonly string[] = []): Uint8Array =>
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
                  .map((c, j) => cell(blank.includes(`${i},${j}`) ? '' : c))
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
