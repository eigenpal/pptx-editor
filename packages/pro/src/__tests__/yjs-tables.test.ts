import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import {
  createPptxEditor,
  type EditorCollaborationSession,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  textOf,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

// Table editing between two editors on two linked Y.Docs (AGENTS.md, "Collaboration
// verification"): concurrent cell typing and row inserts, undo and redo of a merge, and a row
// or table deleted while the other participant types in it. Both replicas converge and the
// result survives save and reopen.

const IN = 914_400;
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
          cy: 2 * 370_840,
          columns: [IN, IN, IN],
          tblPr: '<a:tblPr firstRow="1" bandRow="1"/>',
          rows: ['ABC', 'DEF']
            .map((r) => `<a:tr h="370840">${[...r].map((c) => cell(c)).join('')}</a:tr>`)
            .join(''),
        }),
      },
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
function linkDocs(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

function linkAwareness(a: Awareness, b: Awareness): void {
  const send =
    (from: Awareness, to: Awareness) =>
    (
      { added, updated, removed }: { added: number[]; updated: number[]; removed: number[] },
      origin: unknown
    ) => {
      if (origin === 'link') return;
      applyAwarenessUpdate(
        to,
        encodeAwarenessUpdate(from, [...added, ...updated, ...removed]),
        'link'
      );
    };
  a.on('update', send(a, b));
  b.on('update', send(b, a));
}

async function replica(
  doc: Y.Doc,
  actorId: string,
  awareness: Awareness
): Promise<{ editor: PptxEditorInstance; session: EditorCollaborationSession }> {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId.toUpperCase() },
    awareness,
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(deck())).ok).toBe(true);
  return { editor, session };
}

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = linkDocs(docA, docB);
  const awA = new Awareness(docA);
  const awB = new Awareness(docB);
  linkAwareness(awA, awB);
  const a = await replica(docA, 'a', awA);
  const b = await replica(docB, 'b', awB);
  return { a, b, sync };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const frame = (e: PptxEditorInstance): Element | undefined => {
  let out: Element | undefined;
  const walk = (el: Element) => {
    for (const c of el.children)
      if (c.type === 'element') {
        if (c.kind === 'graphicFrame') out = c;
        else walk(c);
      }
  };
  walk(root(e));
  return out;
};
/** The text of every cell, row by row. */
const grid = (e: PptxEditorInstance): string[] => {
  const tbl = frame(e);
  if (!tbl) return [];
  const rows: string[] = [];
  const walk = (el: Element) => {
    for (const c of el.children)
      if (c.type === 'element') {
        if (c.local === 'tr')
          rows.push(
            c.children
              .filter((tc): tc is Element => tc.type === 'element' && tc.local === 'tc')
              .map((tc) => textOf(tc))
              .join('|')
          );
        else walk(c);
      }
  };
  walk(tbl);
  return rows;
};

function caret(e: PptxEditorInstance, row: number, column: number, offset: number): void {
  e.setSelection({
    kind: 'text',
    nodeId: frame(e)!.id,
    cell: { row, column },
    anchor: { paragraph: 0, offset },
    focus: { paragraph: 0, offset },
  });
}

/**
 * Select the first letter of a cell. A table edit there commits at once: Bold at a bare caret on
 * a word edge only arms for the next typed text (caret-arm.ts), and arming is no edit.
 */
function firstLetter(e: PptxEditorInstance, row: number, column: number): void {
  e.setSelection({
    kind: 'text',
    nodeId: frame(e)!.id,
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 1 },
  });
}

/**
 * Both replicas hold the same deck, and it survives save and reopen. A save repairs a table
 * whose structure peers changed at once, so both saves, not the store, are compared.
 */
function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  const saved = (e: PptxEditorInstance) => {
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    return r.ok ? semanticDigest(r.pkg) : '';
  };
  expect(saved(a)).toBe(saved(b));
}

describe('tables in a collaboration room', () => {
  test('typing in a cell and inserting a row at the same time both survive', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 0, 0, 1);
    expect(a.editor.input('text', '1').ok).toBe(true);
    caret(b.editor, 1, 2, 1);
    expect(b.editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    sync.flush();
    expect(grid(a.editor)).toEqual(['A1|B|C', 'D|E|F', '||']);
    converged(a.editor, b.editor);
  });

  test('a merge, its undo, and its redo reach the other replica', async () => {
    const { a, b } = await pair();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    // The covered cell keeps one empty paragraph, as outside a room.
    expect(grid(b.editor)[0]).toBe('AB||C');
    converged(a.editor, b.editor);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(grid(b.editor)[0]).toBe('A|B|C');
    expect(a.editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(grid(b.editor)[0]).toBe('AB||C');
    converged(a.editor, b.editor);
  });

  test('a row deleted while the peer types in it: both replicas end the same', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 1, 1, 1);
    expect(a.editor.input('text', 'x').ok).toBe(true);
    caret(b.editor, 1, 0, 0);
    expect(b.editor.exec({ type: 'deleteRow' }).ok).toBe(true);
    sync.flush();
    converged(a.editor, b.editor);
    expect(grid(a.editor)).toEqual(grid(b.editor));
    expect(grid(a.editor)[0]).toBe('A|B|C');
  });

  test('the table deleted while the peer types in it: the caret falls back', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 0, 1, 1);
    a.editor.input('text', 'y');
    b.editor.setSelection({ kind: 'shapes', nodeIds: [frame(b.editor)!.id] });
    expect(b.editor.exec({ type: 'deleteTable' }).ok).toBe(true);
    sync.flush();
    converged(a.editor, b.editor);
    expect(!!frame(a.editor)).toBe(!!frame(b.editor));
    if (!frame(a.editor)) expect(a.editor.snapshot().selection).toEqual({ kind: 'none' });
  });

  test("a caret in a cell reaches the peer as text in that cell's body", async () => {
    const { a, b } = await pair();
    caret(a.editor, 1, 2, 1);
    const remote = b.session.remoteSelections!()[0]?.selection;
    expect(remote?.kind).toBe('text');
    // The point is the text node of cell (1, 2) on B: "F", offset 1.
    const node = remote?.kind === 'text' ? remote.anchor.node : '';
    let hit: string | undefined;
    const walk = (el: Element) => {
      for (const c of el.children) {
        if (c.type === 'text' && c.id === node) hit = c.value;
        if (c.type === 'element') walk(c);
      }
    };
    walk(root(b.editor));
    expect(hit).toBe('F');
  });
});

/** The `a:tr` rows of a frame, each as its `a:tc` cells. */
function rowsOf(tbl: Element): Element[][] {
  const out: Element[][] = [];
  const walk = (el: Element) => {
    for (const c of el.children)
      if (c.type === 'element') {
        if (c.local === 'tr')
          out.push(
            c.children.filter((x): x is Element => x.type === 'element' && x.local === 'tc')
          );
        else walk(c);
      }
  };
  walk(tbl);
  return out;
}

/**
 * The deck a replica saves holds a table that opens without repair: one `a:tc` per
 * `a:gridCol` in every row, and every merge marker on a cell that a span covers (and no
 * covered cell without one).
 */
function savedTableValid(e: PptxEditorInstance): void {
  const r = readPptx(e.save()!);
  expect(r.ok).toBe(true);
  if (!r.ok) return;
  const slide = xmlPartOf(r.pkg, e.session!.info.slides[0].partName)!.root;
  let tbl: Element | undefined;
  const find = (el: Element) => {
    for (const c of el.children) if (c.type === 'element') c.local === 'tbl' ? (tbl = c) : find(c);
  };
  find(slide);
  let columns = 0;
  const countCols = (el: Element) => {
    for (const c of el.children)
      if (c.type === 'element') c.local === 'gridCol' ? (columns += 1) : countCols(c);
  };
  countCols(tbl!);
  const rows = rowsOf(tbl!);
  expect(rows.map((r) => r.length)).toEqual(rows.map(() => columns));
  const at = (tc: Element, name: string) => tc.attrs.find((x) => x.local === name)?.value;
  const covered = rows.map((r) => r.map(() => false));
  rows.forEach((row, r) =>
    row.forEach((tc, c) => {
      if (covered[r][c] || at(tc, 'hMerge') || at(tc, 'vMerge')) return;
      const g = Math.min(Number(at(tc, 'gridSpan') ?? 1), columns - c);
      const k = Math.min(Number(at(tc, 'rowSpan') ?? 1), rows.length - r);
      for (let dr = 0; dr < k; dr += 1)
        for (let dc = 0; dc < g; dc += 1) if (dr || dc) covered[r + dr][c + dc] = true;
    })
  );
  const marked = rows.map((row) => row.map((tc) => !!(at(tc, 'hMerge') || at(tc, 'vMerge'))));
  expect(marked).toEqual(covered);
}

describe('concurrent table structure edits', () => {
  test('a row and a column inserted at the same time give a valid table', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 0, 0, 0);
    expect(a.editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    caret(b.editor, 0, 1, 0);
    expect(b.editor.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    // Each replica saves a repaired copy; the stores stay equal.
    savedTableValid(a.editor);
    savedTableValid(b.editor);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    // The next table edit repairs the table in the room, and both replicas follow.
    firstLetter(a.editor, 0, 0);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(rowsOf(frame(a.editor)!).map((r) => r.length)).toEqual([4, 4, 4]);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    converged(a.editor, b.editor);
  });

  test('two peers that repair the same table at the same time converge on one grid', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 0, 0, 0);
    expect(a.editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    caret(b.editor, 0, 1, 0);
    expect(b.editor.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    sync.flush();
    // Both peers repair the short row in their next edit, at the same time.
    sync.pause();
    firstLetter(a.editor, 0, 0);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    caret(b.editor, 2, 0, 0);
    expect(b.editor.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    // The next edit trims the blank cell the second repair added; no grid column is added.
    caret(a.editor, 1, 1, 0);
    expect(a.editor.exec({ type: 'toggleUnderline' }).ok).toBe(true);
    expect(rowsOf(frame(a.editor)!).map((r) => r.length)).toEqual([4, 4, 4]);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    savedTableValid(b.editor);
  });

  test('a column deleted while the peer inserts a row gives a valid table', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 1, 0, 0);
    expect(a.editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    caret(b.editor, 0, 2, 0);
    expect(b.editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    savedTableValid(a.editor);
    savedTableValid(b.editor);
  });

  test('a merge beside a concurrent column insert gives valid merge markers', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    caret(b.editor, 0, 0, 0);
    expect(b.editor.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    // The next table edit rewrites the merge markers from the spans, for both peers.
    firstLetter(a.editor, 1, 3);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    savedTableValid(a.editor);
    savedTableValid(b.editor);
  });

  test('a merge beside a concurrent column delete leaves the next cell visible', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    caret(b.editor, 1, 1, 0);
    expect(b.editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    // The merge's span now reaches "C", which was never merged. The next table edit ends the
    // span before it, so "C" shows again, for both peers.
    firstLetter(a.editor, 1, 0);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    const row0 = rowsOf(frame(a.editor)!)[0];
    expect(row0.map((tc) => textOf(tc))).toEqual(['AB', 'C']);
    expect(row0.map((tc) => tc.attrs.map((x) => x.local).join(','))).toEqual(['', '']);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    savedTableValid(b.editor);
  });

  test('a merge and a concurrent delete of the column of its top-left cell keep the merged text', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    caret(b.editor, 1, 0, 0);
    expect(b.editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    // "A" went with its column, as the peer asked. "B" stays: it goes back to its cell, which
    // no span covers any more, and the next table command writes that cell as a plain cell.
    firstLetter(b.editor, 1, 0);
    expect(b.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    sync.flush();
    const row0 = rowsOf(frame(a.editor)!)[0];
    expect(row0.map((tc) => textOf(tc))).toEqual(['B', 'C']);
    expect(row0.map((tc) => tc.attrs.map((x) => x.local).join(','))).toEqual(['', '']);
    converged(a.editor, b.editor);
    savedTableValid(a.editor);
  });

  test('a vertical merge and a concurrent delete of the row of its top-left cell keep the merged text', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 0 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    caret(b.editor, 0, 2, 0);
    expect(b.editor.exec({ type: 'deleteRow' }).ok).toBe(true);
    sync.flush();
    firstLetter(a.editor, 0, 1);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    sync.flush();
    expect(grid(a.editor)).toEqual(['D|E|F']);
    converged(a.editor, b.editor);
    savedTableValid(b.editor);
  });

  test('a vertical merge and a concurrent delete of its first row draw the kept text at once', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 0 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    caret(b.editor, 0, 2, 0);
    expect(b.editor.exec({ type: 'deleteRow' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    // The cell of "D" keeps vMerge, and no span reaches it any more: both replicas draw it with
    // its text before any table command, and find reads it.
    for (const e of [a.editor, b.editor]) {
      const frameItem = e.session!.layout(0)!.items.find((i) => i.kind === 'frame');
      const cells = frameItem?.kind === 'frame' ? (frameItem.table?.cells ?? []) : [];
      expect(cells.map((c) => `${c.cell.row},${c.cell.col}`)).toEqual(['0,0', '0,1', '0,2']);
      expect(e.find.matches('D')).toHaveLength(1);
    }
    // The cell takes typing, and both replicas converge on it.
    caret(b.editor, 0, 0, 0);
    expect(b.editor.input('text', 'x').ok).toBe(true);
    sync.flush();
    expect(grid(a.editor)[0].split('|')[0]).toBe('xD');
    converged(a.editor, b.editor);
    // At the same time, A makes the orphan cell's text bold and B inserts a row. Both commands
    // repair the table's merge markers from the drawn cells, so the repairs agree.
    sync.pause();
    firstLetter(a.editor, 0, 0);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    caret(b.editor, 0, 2, 0);
    expect(b.editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    sync.flush();
    converged(a.editor, b.editor);
    expect(grid(a.editor)[0].split('|')[0]).toBe('xD');
    savedTableValid(a.editor);
  });

  test('text deleted from a merged cell is not saved', async () => {
    const { a } = await pair();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    // Delete the merged text, then save: no cell keeps a hidden "B".
    a.editor.setSelection({
      kind: 'text',
      nodeId: frame(a.editor)!.id,
      cell: { row: 0, column: 0 },
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 1, offset: 1 },
    });
    expect(a.editor.input('deleteBackward').ok).toBe(true);
    expect(grid(a.editor)[0]).toBe('||C');
    const reopened = readPptx(a.editor.save()!);
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) return;
    const slide = xmlPartOf(reopened.pkg, a.editor.session!.info.slides[0].partName)!.root;
    expect(textOf(slide)).not.toContain('B');
  });

  test('typing in a merged cell while the peer deletes its first column keeps the typing', async () => {
    const { a, b, sync } = await pair();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    sync.pause();
    caret(b.editor, 0, 0, 0);
    expect(b.editor.input('text', 'x').ok).toBe(true);
    caret(a.editor, 1, 0, 0);
    expect(a.editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
    sync.flush();
    expect(grid(a.editor)[0].split('|')[0]).toBe('xAB');
    converged(a.editor, b.editor);
    savedTableValid(a.editor);
  });

  test('two overlapping merges at the same time converge, and the next table edit shows every text', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    b.editor.setSelection({
      kind: 'cells',
      nodeId: frame(b.editor)!.id,
      anchor: { row: 0, column: 1 },
      focus: { row: 0, column: 2 },
    });
    expect(b.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    firstLetter(a.editor, 1, 0);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    sync.flush();
    const shown = a.editor.find.matches('C');
    expect(shown.length).toBeGreaterThan(0);
    converged(a.editor, b.editor);
    savedTableValid(a.editor);
  });

  test('a pasted copy of a merged table holds no mark', async () => {
    const { a } = await pair();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    a.editor.setSelection({ kind: 'shapes', nodeIds: [frame(a.editor)!.id] });
    expect(a.editor.exec({ type: 'paste', payload: a.editor.copy()! }).ok).toBe(true);
    const tables: Element[] = [];
    const walk = (el: Element) => {
      for (const c of el.children)
        if (c.type === 'element') c.local === 'tbl' ? tables.push(c) : walk(c);
    };
    walk(root(a.editor));
    expect(tables).toHaveLength(2);
    expect(rowsOf(tables[1])[0].map((tc) => textOf(tc))).toEqual(['AB', '', 'C']);
    expect(JSON.stringify(tables[1])).not.toContain('table-structure-edit');
  });

  test('typing in a cell while the peer merges it reaches the merged cell', async () => {
    const { a, b, sync } = await pair();
    caret(b.editor, 0, 1, 1);
    sync.pause();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    expect(b.editor.input('text', 'x').ok).toBe(true);
    sync.flush();
    expect(grid(a.editor)[0].split('|')[0]).toBe('ABx');
    converged(a.editor, b.editor);
  });

  test('merged text is found once, saved once, and stays in the merged cell after a split', async () => {
    const { a, b } = await pair();
    a.editor.setSelection({
      kind: 'cells',
      nodeId: frame(a.editor)!.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(a.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    expect(b.editor.find.matches('B')).toHaveLength(1);
    const reopened = readPptx(a.editor.save()!);
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) return;
    const slide = xmlPartOf(reopened.pkg, a.editor.session!.info.slides[0].partName)!.root;
    let tbl: Element | undefined;
    const walk = (el: Element) => {
      for (const c of el.children)
        if (c.type === 'element') c.local === 'tbl' ? (tbl = c) : walk(c);
    };
    walk(slide);
    expect(rowsOf(tbl!)[0].map((tc) => textOf(tc))).toEqual(['AB', '', 'C']);
    // The structure mark never reaches the file.
    expect(new TextDecoder().decode(a.editor.save()!)).not.toContain('table-structure-edit');
    caret(b.editor, 0, 0, 0);
    expect(b.editor.exec({ type: 'splitCell', rows: 1, cols: 2 }).ok).toBe(true);
    expect(grid(a.editor)[0]).toBe('AB||C');
    converged(a.editor, b.editor);
  });

  test('a caret in a cell the peer merges away moves to the merged cell', async () => {
    const { a, b, sync } = await pair();
    caret(a.editor, 1, 1, 0);
    sync.pause();
    b.editor.setSelection({
      kind: 'cells',
      nodeId: frame(b.editor)!.id,
      anchor: { row: 1, column: 0 },
      focus: { row: 1, column: 1 },
    });
    expect(b.editor.exec({ type: 'mergeCells' }).ok).toBe(true);
    sync.flush();
    const sel = a.editor.snapshot().selection;
    expect(sel.kind === 'text' && sel.cell).toEqual({ row: 1, column: 0 });
    expect(a.editor.input('text', 'z').ok).toBe(true);
    sync.flush();
    expect(grid(a.editor)[1].split('|')[0]).toContain('z');
    converged(a.editor, b.editor);
  });
});

describe('concurrent table formatting', () => {
  test('borders set on two neighbor cells at once leave one line per edge', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    caret(a.editor, 0, 0, 0);
    const spec = { color: '#C00000', width: 2 } as const;
    expect(a.editor.exec({ type: 'setTableBorders', scope: 'right', spec }).ok).toBe(true);
    caret(b.editor, 0, 1, 0);
    expect(b.editor.exec({ type: 'setTableBorders', scope: 'left', spec }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    const lines = (e: PptxEditorInstance, bytes?: Uint8Array) => {
      const read = bytes ? readPptx(bytes) : null;
      if (read && !read.ok) throw new Error('reopen');
      const pkg = read ? read.pkg : e.session!.store.package;
      const slide = xmlPartOf(pkg, e.session!.info.slides[0].partName)!.root;
      let tcPr: Element | undefined;
      const walk = (el: Element) => {
        for (const c of el.children)
          if (c.type === 'element') c.local === 'tcPr' && !tcPr ? (tcPr = c) : walk(c);
      };
      walk(slide);
      return tcPr!.children.filter((c): c is Element => c.type === 'element').map((c) => c.local);
    };
    // The saved copy holds one lnR in cell (0, 0); the next table edit repairs the room too.
    expect(lines(a.editor, a.editor.save()!).filter((n) => n === 'lnR')).toHaveLength(1);
    firstLetter(a.editor, 1, 1);
    expect(a.editor.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(lines(a.editor).filter((n) => n === 'lnR')).toHaveLength(1);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
  });
});

// The text format commands act on table cells too: each runs on a cell while the other
// participant changes the table's structure or formats the next cell.
describe('text formats in table cells', () => {
  const xml = (e: PptxEditorInstance) => JSON.stringify(frame(e));

  test('Font, Paragraph, and Bullets on a cell while the peer inserts a row converge', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    firstLetter(a.editor, 0, 1);
    expect(a.editor.exec({ type: 'setFontFormat', caps: 'small', characterSpacingPt: 2 }).ok).toBe(
      true
    );
    expect(a.editor.exec({ type: 'setParagraphFormat', spaceBeforePt: 6 }).ok).toBe(true);
    const numbers = { kind: 'number', scheme: 'arabicPeriod' } as const;
    expect(a.editor.exec({ type: 'setBulletFormat', format: numbers }).ok).toBe(true);
    caret(b.editor, 0, 0, 0);
    expect(b.editor.exec({ type: 'insertRow', where: 'above' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    expect(grid(a.editor)).toEqual(['||', 'A|B|C', 'D|E|F']);
    for (const e of [a.editor, b.editor]) {
      const text = xml(e);
      expect(text).toContain('"spc"');
      expect(text).toContain('buAutoNum');
      savedTableValid(e);
    }
    converged(a.editor, b.editor);
  });

  test('two peers format neighbor cells at once; undo removes only its own format', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    firstLetter(a.editor, 1, 0);
    expect(a.editor.exec({ type: 'toggleSuperscript' }).ok).toBe(true);
    firstLetter(b.editor, 1, 1);
    expect(b.editor.exec({ type: 'changeCase', case: 'lower' }).ok).toBe(true);
    sync.flush();
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    expect(grid(a.editor)).toEqual(['A|B|C', 'D|e|F']);
    expect(xml(a.editor)).toContain('"baseline"');
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(xml(b.editor)).not.toContain('"baseline"');
    expect(grid(b.editor)).toEqual(['A|B|C', 'D|e|F']);
    converged(a.editor, b.editor);
  });
});
