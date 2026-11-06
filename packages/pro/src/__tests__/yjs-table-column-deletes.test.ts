/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Concurrent deletes of the last columns of a table after a row insert
// (AGENTS.md, "Collaboration verification"). Three editors on linked Y.Docs open a 3 x 3 table;
// peer C deletes column 2, then, with the link paused, peer A inserts a row and deletes one of
// the two columns left while peer C deletes the other (or the same) one. The inserted row keeps
// its cell of the column that peer C deleted, because peer C never saw it. When no column is
// left, the table goes on save and on the canvas, as a local delete of the last column does.
// Two peers who delete the last two rows, one each, leave no row: the table goes too.
// Each case runs in both client orders, with undo and redo. Every replica saves with no local
// write, the saves reopen to the same valid package, and the canvas grid equals the saved grid.

import { describe, expect, test } from 'bun:test';
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
import { IN, all, kids, peers, root, slide1 } from './lineage-support.ts';

const deck = (): Uint8Array =>
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
            .map((r) => `<a:tr h="370840">${[...r].map((c) => cell(c)).join('')}</a:tr>`)
            .join(''),
        }),
      },
    ],
  });

/** The slide's graphic frame (PresentationML, so `all` does not find it). */
function frameOf(e: PptxEditorInstance): Element | undefined {
  const find = (el: Element): Element | undefined => {
    for (const c of kids(el)) {
      const hit = c.kind === 'graphicFrame' ? c : find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  return find(root(e));
}

/** Run a row or column command on `e` from a caret in cell (`row`, `column`). */
function exec(
  e: PptxEditorInstance,
  row: number,
  column: number,
  command: { type: 'deleteColumn' | 'deleteRow' } | { type: 'insertRow'; where: 'above' }
): void {
  e.setSelection({
    kind: 'text',
    nodeId: frameOf(e)!.id,
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
  expect(e.exec(command).ok).toBe(true);
}

/**
 * The table that `e` draws: per row, the text of each drawn cell. [] when it draws none. A
 * blank cell that the save adds to a short row has no node in the store, and no text.
 */
function drawn(e: PptxEditorInstance): string[] {
  const item = e.session!.layout(0)?.items.find((i) => i.element.kind === 'graphicFrame');
  const t = item?.element.kind === 'graphicFrame' ? item.element.table : undefined;
  if (!t) return [];
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out: string[] = [];
  for (const c of t.cells) {
    if (c.col === 0) out.push('');
    const tc = byId.get(c.nodeId);
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : tc ? textOf(tc) : ''}`;
  }
  return out;
}

/**
 * Every replica saves with no local write before it. The saves reopen to the same package
 * (`canonicalPackageFingerprint` and `semanticDigest`) and pass the part checks, and each
 * replica draws the grid of its saved deck. Returns the saved table's rows, or null when the
 * save leaves the table out.
 */
async function savedEqual(editors: readonly PptxEditorInstance[]): Promise<string[] | null> {
  const [first, ...rest] = editors;
  const print = canonicalPackageFingerprint(first!.session!.store.package);
  for (const e of rest) expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(print);
  const revisions = editors.map((e) => e.session!.store.revision);
  const saves = editors.map((e) => e.save()!);
  expect(editors.map((e) => e.session!.store.revision)).toEqual(revisions);
  const reads = saves.map((bytes) => readPptx(bytes));
  const pkgs = reads.map((r) => (r.ok ? r.pkg : undefined));
  const r = { pkg: pkgs[0]! };
  for (const pkg of pkgs.slice(1)) {
    expect(canonicalPackageFingerprint(pkg!)).toBe(canonicalPackageFingerprint(r.pkg));
    expect(semanticDigest(pkg!)).toBe(semanticDigest(r.pkg));
  }
  expect(packageViolations(r.pkg)).toEqual([]);
  for (const [i, e] of editors.entries()) {
    const reopened = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await reopened.load(saves[i]!)).ok).toBe(true);
    expect(drawn(e)).toEqual(drawn(reopened));
    reopened.destroy();
  }
  const tbl = all(xmlPartOf(r.pkg, slide1(first!))!.root, 'tbl')[0];
  if (!tbl) return null;
  expect(all(tbl, 'tr').every((tr) => all(tr, 'tc').length === all(tbl, 'gridCol').length)).toBe(
    true
  );
  return all(tbl, 'tr').map((tr) =>
    kids(tr)
      .filter((c) => c.local === 'tc')
      .map(textOf)
      .join('|')
  );
}

/** The client-id orders: peer A's id below peer C's, then above it. */
const ORDERS = [
  { name: 'A before C', ids: [1_001, 2_002, 3_003] },
  { name: 'C before A', ids: [3_003, 2_002, 1_001] },
] as const;

/**
 * One history: C deletes column 2 and all updates arrive; then, with the link paused,
 * A inserts a row above row 3 and deletes column 2 of the two left, and C deletes column
 * `cColumn` (1-based, of the two left). The link then flushes.
 */
async function history(ids: readonly number[], cColumn: 1 | 2) {
  const { editors, sync } = await peers(deck(), ids);
  const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
  exec(c, 0, 1, { type: 'deleteColumn' });
  sync.flush();
  sync.pause();
  exec(a, 2, 0, { type: 'insertRow', where: 'above' });
  exec(a, 0, 1, { type: 'deleteColumn' });
  exec(c, 1, cColumn - 1, { type: 'deleteColumn' });
  sync.flush();
  return { a, b, c, editors, sync };
}

describe('concurrent deletes of the last columns after a row insert', () => {
  for (const order of ORDERS)
    test(`different columns: the table goes on save and on the canvas (${order.name})`, async () => {
      const { a, b, c, editors } = await history(order.ids, 1);
      expect(await savedEqual(editors)).toBeNull();
      for (const e of [a, b, c]) {
        expect(drawn(e)).toEqual([]);
        // The room keeps the frame: an undo can bring a column back. The selection pane does
        // not list it.
        expect(frameOf(e)).toBeDefined();
        expect(e.snapshot().shapes).toEqual([]);
      }
      // The carets that A and C left in the table reach no drawn cell: typing and a table
      // command there are refused and write nothing.
      for (const e of [a, c]) {
        const revision = e.session!.store.revision;
        expect(e.input('text', 'x').ok).toBe(false);
        expect(e.exec({ type: 'deleteRow' }).ok).toBe(false);
        expect(e.session!.store.revision).toBe(revision);
      }
    });

  for (const order of ORDERS)
    test(`different columns, undo and redo on each side (${order.name})`, async () => {
      const { a, c, editors, sync } = await history(order.ids, 1);
      // C's undo brings column 1 back, with the inserted row's cell that C never removed.
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(await savedEqual(editors)).toEqual(['A', 'D', '', 'G']);
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      sync.flush();
      expect(await savedEqual(editors)).toBeNull();
      // A's undo brings column 2 back, in every row.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      const back = await savedEqual(editors);
      expect(back).not.toBeNull();
      expect(back!.map((row) => row.replace(/\|/g, ''))).toEqual(['C', 'F', '', 'I']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      sync.flush();
      expect(await savedEqual(editors)).toBeNull();
    });

  for (const order of ORDERS)
    test(`the same column: the other column stays, with the inserted row (${order.name})`, async () => {
      const { c, editors, sync } = await history(order.ids, 2);
      expect(await savedEqual(editors)).toEqual(['A', 'D', '', 'G']);
      // An undo of either delete brings the column back in the rows that its peer knew, and
      // leaves the inserted row one cell short. The save adds the cell, and the canvas draws it.
      expect(c.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(await savedEqual(editors)).toEqual(['A|C', 'D|F', '|', 'G|I']);
      expect(c.exec({ type: 'redo' }).ok).toBe(true);
      sync.flush();
      expect(await savedEqual(editors)).toEqual(['A', 'D', '', 'G']);
    });

  for (const order of ORDERS)
    test(`the last two rows, one by each peer: the table goes, and an undo brings it back (${order.name})`, async () => {
      const { editors, sync } = await peers(deck(), order.ids);
      const [a, , c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      exec(c, 2, 0, { type: 'deleteRow' });
      sync.flush();
      sync.pause();
      exec(a, 0, 0, { type: 'deleteRow' });
      exec(c, 1, 0, { type: 'deleteRow' });
      sync.flush();
      expect(await savedEqual(editors)).toBeNull();
      expect(drawn(a)).toEqual([]);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(await savedEqual(editors)).toEqual(['A|B|C']);
    });
});
