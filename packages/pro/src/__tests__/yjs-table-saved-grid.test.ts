/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Concurrent table structure edits that leave rows with too few or too many cells
// (eigenpal/pptx-editor-miro#143; AGENTS.md, "Collaboration verification"). Three editors on linked
// Y.Docs open a 3 x 3 table and, with the link paused, change its structure at the same time: a
// row insert and a column insert leave the new row one cell short, two splits in one row leave
// the rows uneven, and a column delete beside a split leaves a row one cell long. The save
// repairs the cell counts, and the canvas draws the table as the save writes it: a short row
// shows the blank cells the save adds, and a long row shows the grid column the save adds.
// Each history runs in both client orders, with undo and redo. On every replica, the canvas
// grid equals the grid of the saved deck reopened, the saves reopen to the same valid package,
// and text typed into each drawn cell shows in that cell on every replica and in the save.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  textOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import { IN, all, kids, peers, root } from './lineage-support.ts';

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

/** Put a caret at offset `offset` of cell (`row`, `column`) of `e`'s table. */
function caret(e: PptxEditorInstance, row: number, column: number, offset = 0): void {
  e.setSelection({
    kind: 'text',
    nodeId: frameOf(e)!.id,
    cell: { row, column },
    anchor: { paragraph: 0, offset },
    focus: { paragraph: 0, offset },
  });
}

type Structure =
  | { type: 'insertRow'; where: 'below' }
  | { type: 'insertColumn'; where: 'right' }
  | { type: 'deleteColumn' }
  | { type: 'splitCell'; rows: number; cols: number };

/** Run a structure command on `e` from a caret in cell (`row`, `column`). */
function exec(e: PptxEditorInstance, row: number, column: number, command: Structure): void {
  caret(e, row, column);
  expect(e.exec(command).ok).toBe(true);
}

/** The laid-out table of `e`. */
function layoutTable(e: PptxEditorInstance) {
  const item = e.session!.layout(0)?.items.find((i) => i.element.kind === 'graphicFrame');
  return item?.element.kind === 'graphicFrame' ? item.element.table : undefined;
}

/**
 * The table that `e` draws: per row, each drawn cell as its text with its spans when they are
 * not 1 x 1, and `·` for a covered position. A blank cell that the save adds to a short row has
 * no node in the store yet, so it draws no text.
 */
function drawn(e: PptxEditorInstance): string[] {
  const t = layoutTable(e);
  if (!t) return [];
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out: string[] = [];
  for (const c of t.cells) {
    if (c.col === 0) out.push('');
    const tc = byId.get(c.nodeId);
    expect(tc !== undefined || c.nodeId === '').toBe(true);
    const spans = c.gridSpan > 1 || c.rowSpan > 1 ? `[${c.rowSpan}x${c.gridSpan}]` : '';
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : (tc ? textOf(tc) : '') + spans}`;
  }
  return out;
}

/**
 * The replicas hold the same package; every replica saves with no local write; the saves
 * reopen to the same valid package, one `a:tc` per `a:gridCol` in every row; and each replica
 * draws the grid of its saved deck. Returns that grid.
 */
async function savedEqual(editors: readonly PptxEditorInstance[]): Promise<string[]> {
  const print = canonicalPackageFingerprint(editors[0]!.session!.store.package);
  for (const e of editors)
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(print);
  const revisions = editors.map((e) => e.session!.store.revision);
  const saves = editors.map((e) => e.save()!);
  expect(editors.map((e) => e.session!.store.revision)).toEqual(revisions);
  const pkgs = saves.map((bytes) => {
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('the save does not open');
    return r.pkg;
  });
  for (const pkg of pkgs) {
    expect(canonicalPackageFingerprint(pkg)).toBe(canonicalPackageFingerprint(pkgs[0]!));
    expect(semanticDigest(pkg)).toBe(semanticDigest(pkgs[0]!));
  }
  expect(packageViolations(pkgs[0]!)).toEqual([]);
  let grid: string[] = [];
  for (const [i, e] of editors.entries()) {
    const reopened = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await reopened.load(saves[i]!)).ok).toBe(true);
    const tbl = all(root(reopened), 'tbl')[0]!;
    const columns = all(tbl, 'gridCol').length;
    expect(
      all(tbl, 'tr').every((tr) => kids(tr).filter((c) => c.local === 'tc').length === columns)
    ).toBe(true);
    grid = drawn(reopened);
    expect(drawn(e)).toEqual(grid);
    reopened.destroy();
  }
  return grid;
}

/** The client-id orders: peer A's id below peer B's, then above it. */
const ORDERS = [
  { name: 'A before B', ids: [1_001, 2_002, 3_003] },
  { name: 'B before A', ids: [3_003, 2_002, 1_001] },
] as const;

/** One history of eigenpal/pptx-editor-miro#143: two concurrent structure edits, on peers A and B. */
interface History {
  readonly name: string;
  readonly a: { row: number; column: number; command: Structure };
  readonly b: { row: number; column: number; command: Structure };
  /** The grid that every replica draws and saves after the flush. */
  readonly grid: readonly string[];
}

const HISTORIES: readonly History[] = [
  {
    name: 'a row insert and a column insert',
    a: { row: 2, column: 2, command: { type: 'insertRow', where: 'below' } },
    b: { row: 2, column: 2, command: { type: 'insertColumn', where: 'right' } },
    grid: ['A|B|C|', 'D|E|F|', 'G|H|I|', '|||'],
  },
  {
    name: 'two splits in one row',
    a: { row: 1, column: 1, command: { type: 'splitCell', rows: 1, cols: 2 } },
    b: { row: 1, column: 2, command: { type: 'splitCell', rows: 2, cols: 1 } },
    // The split into two rows adds a row with one cell for each of the three columns it knew;
    // the other split adds a fourth column. The save gives the new row its fourth cell.
    grid: ['A|B[1x2]|·|C', 'D[2x1]|E[2x1]||F', '·|·||', 'G|H[1x2]|·|I'],
  },
  {
    name: 'a column delete and a split of a cell beside it',
    a: { row: 0, column: 0, command: { type: 'splitCell', rows: 2, cols: 2 } },
    b: { row: 2, column: 1, command: { type: 'deleteColumn' } },
    // The split's new row keeps a cell of the deleted column, one more than the grid: the save
    // adds a grid column, and the other rows get a blank cell for it.
    grid: ['A||C[2x1]|', '||·|', 'D[1x2]|·|F|', 'G[1x2]|·|I|'],
  },
];

/** Open three peers, run `h` on A and B with the link paused, and flush. */
async function run(h: History, ids: readonly number[]) {
  const { editors, sync } = await peers(deck(), ids);
  const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
  sync.pause();
  exec(a, h.a.row, h.a.column, h.a.command);
  exec(b, h.b.row, h.b.column, h.b.command);
  sync.flush();
  return { a, b, editors, sync };
}

describe('the canvas draws the grid that the save writes after concurrent structure edits (eigenpal/pptx-editor-miro#143)', () => {
  for (const h of HISTORIES)
    for (const order of ORDERS) {
      test(`${h.name} (${order.name})`, async () => {
        const { editors } = await run(h, order.ids);
        expect(await savedEqual(editors)).toEqual([...h.grid]);
      });

      test(`${h.name}: typing into every drawn cell (${order.name})`, async () => {
        const { editors, sync } = await run(h, order.ids);
        const t = layoutTable(editors[0]!)!;
        const cells = t.cells.filter((c) => !c.covered);
        // Each peer types into one cell in turn, so the peers that did not change the
        // structure type too.
        cells.forEach((c, i) => {
          const e = editors[i % editors.length]!;
          caret(e, c.row, c.col);
          expect(e.input('text', String.fromCharCode(0x100 + i)).ok).toBe(true);
          sync.flush();
        });
        // Each letter shows in the cell it was typed in, and the grid keeps its shape.
        const grid = (await savedEqual(editors)).map((row) => row.split('|'));
        cells.forEach((c, i) =>
          expect(grid[c.row]![c.col]).toContain(String.fromCharCode(0x100 + i))
        );
        expect(grid.map((row) => row.length)).toEqual(h.grid.map((row) => row.split('|').length));
      });

      test(`${h.name}: undo and redo on each side (${order.name})`, async () => {
        const { a, b, editors, sync } = await run(h, order.ids);
        for (const e of [a, b]) {
          expect(e.exec({ type: 'undo' }).ok).toBe(true);
          sync.flush();
          await savedEqual(editors);
          expect(e.exec({ type: 'redo' }).ok).toBe(true);
          sync.flush();
          expect(await savedEqual(editors)).toEqual([...h.grid]);
        }
        // Both undone: the table of the file comes back.
        expect(a.exec({ type: 'undo' }).ok).toBe(true);
        expect(b.exec({ type: 'undo' }).ok).toBe(true);
        sync.flush();
        expect(await savedEqual(editors)).toEqual(['A|B|C', 'D|E|F', 'G|H|I']);
      });
    }

  // Two peers that type at once into cells that only the save adds each write those cells: the
  // row then holds more cells than the grid, and the save adds a grid column for them. Both
  // letters stay, the replicas converge, and the canvas draws what the save writes. Where each
  // letter lands follows the client ids, as for any two concurrent inserts at one place.
  for (const order of ORDERS)
    test(`concurrent typing into one added cell keeps both letters (${order.name})`, async () => {
      const { editors, sync } = await run(HISTORIES[0]!, order.ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      caret(a, 3, 3);
      expect(a.input('text', 'x').ok).toBe(true);
      caret(b, 3, 3);
      expect(b.input('text', 'y').ok).toBe(true);
      sync.flush();
      const grid = await savedEqual(editors);
      expect(grid.slice(0, 3)).toEqual(['A|B|C||', 'D|E|F||', 'G|H|I||']);
      expect(['|||x|y', '|||y|x']).toContain(grid[3]!);
    });

  for (const order of ORDERS)
    test(`concurrent typing into two added cells of one row keeps both letters (${order.name})`, async () => {
      const { editors, sync } = await peers(deck(), order.ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      exec(a, 2, 2, { type: 'insertRow', where: 'below' });
      exec(b, 2, 2, { type: 'insertColumn', where: 'right' });
      exec(b, 2, 3, { type: 'insertColumn', where: 'right' });
      sync.flush();
      expect(await savedEqual(editors)).toEqual(['A|B|C||', 'D|E|F||', 'G|H|I||', '||||']);
      sync.pause();
      caret(c, 3, 3);
      expect(c.input('text', 'p').ok).toBe(true);
      caret(a, 3, 4);
      expect(a.input('text', 'q').ok).toBe(true);
      sync.flush();
      const grid = await savedEqual(editors);
      expect(grid.slice(0, 3)).toEqual(['A|B|C|||', 'D|E|F|||', 'G|H|I|||']);
      expect(['||||q|p', '|||p||q']).toContain(grid[3]!);
    });

  test('a cell typed into a short row, then undone, draws blank again (A before B)', async () => {
    const { editors, sync } = await run(HISTORIES[0]!, ORDERS[0].ids);
    const c = editors[2]!;
    caret(c, 3, 3);
    expect(c.input('text', 'x').ok).toBe(true);
    sync.flush();
    expect(await savedEqual(editors)).toEqual(['A|B|C|', 'D|E|F|', 'G|H|I|', '|||x']);
    expect(c.exec({ type: 'undo' }).ok).toBe(true);
    sync.flush();
    expect(await savedEqual(editors)).toEqual([...HISTORIES[0]!.grid]);
    expect(c.exec({ type: 'redo' }).ok).toBe(true);
    sync.flush();
    expect(await savedEqual(editors)).toEqual(['A|B|C|', 'D|E|F|', 'G|H|I|', '|||x']);
  });
});
