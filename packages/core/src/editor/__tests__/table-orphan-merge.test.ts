// Tables whose merge markup disagrees with itself (deleted archive issue 92). Coverage comes from
// spans only (`store/package/table-grid.ts`): a cell with an orphan `hMerge` or `vMerge` (no
// span reaches it) is drawn with its own text and spans, a span past the grid ends at the grid,
// and a span into an earlier block ends before it. The cascade, hit testing, the table commands,
// find, and the store checks read the same grid. A save keeps the file's merge markup unless an
// edit changed the table.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  children,
  elements,
  find,
  readPptx,
  resolveTableGrid,
  textOf,
  xmlPartOf,
  NS,
  type Element,
  type NodeSpec,
  type PptxOp,
} from '../../store/index.ts';
import { TABLE_EDIT_URI } from '../../automation/commands/table-edit-marks.ts';
import { packageViolations } from '../../store/store/part-checks.ts';
import { DEFAULT_TABLE_STYLE_ID } from '../../cascade/index.ts';
import type { EditorCommand } from '../../contracts/index.ts';
import { apply } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { cellAtPoint, tableItem } from '../table-cells.ts';

const IN = 914_400;
const tr = (...cells: string[]): string => `<a:tr h="370840">${cells.join('')}</a:tr>`;

function deck(rows: string, columns: number[]): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 3,
          x: IN,
          y: IN,
          cx: columns.reduce((a, b) => a + b, 0),
          cy: 370_840,
          columns,
          tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
          rows,
        }),
      },
    ],
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const part = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const slideRoot = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, part(e))!.root;
const frameId = (e: PptxEditorInstance): string =>
  elements(find(slideRoot(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
const tcsOf = (root: Element): Element[][] =>
  children(find(root, 'a:tbl'), 'a:tr').map((r) => children(r, 'a:tc'));
const tcs = (e: PptxEditorInstance): Element[][] => tcsOf(slideRoot(e));
const frameOf = (e: PptxEditorInstance) => tableItem(e.session!.layout(0)!, frameId(e))!;

/** The drawn cells: "row,col gridSpan x rowSpan text", in reading order. */
function drawn(e: PptxEditorInstance): string[] {
  const byId = new Map(
    tcs(e)
      .flat()
      .map((tc) => [tc.id, tc])
  );
  return frameOf(e).table.cells.map(
    (c) =>
      `${c.cell.row},${c.cell.col} ${c.cell.gridSpan}x${c.cell.rowSpan} ${textOf(byId.get(c.cell.nodeId))}`
  );
}

/** The merge attributes of every cell, row by row. */
const markup = (root: Element): string[][] =>
  tcsOf(root).map((row) =>
    row.map((tc) =>
      ['gridSpan', 'rowSpan', 'hMerge', 'vMerge']
        .map((n) => (attr(tc, n) === undefined ? '' : `${n}=${attr(tc, n)}`))
        .filter(Boolean)
        .join(' ')
    )
  );

/** The table-merge violations of the slide part. */
const mergeViolations = (e: PptxEditorInstance) =>
  packageViolations(e.session!.store.package).filter((v) => v.rule === 'table-merge');

const caretIn = (e: PptxEditorInstance, row: number, column: number): void =>
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });

/** Save and reopen; the reopened slide part's root. */
function reopened(e: PptxEditorInstance): Element {
  const r = readPptx(e.save()!);
  expect(r.ok).toBe(true);
  if (!r.ok) throw new Error('reopen');
  return xmlPartOf(r.pkg, part(e))!.root;
}

// A 2 x 2 merge whose first row is gone: the rest of a block that a peer's row delete left.
const HEADLESS =
  tr(
    cell('D', { attrs: ' gridSpan="2" vMerge="1"' }),
    cell('', { attrs: ' hMerge="1" vMerge="1"' }),
    cell('F')
  ) + tr(cell('G'), cell('H'), cell('I'));

describe('orphan merge markers are drawn', () => {
  test('a one-cell table whose only cell has vMerge draws the cell and its text', async () => {
    const e = await open(deck(tr(cell('B', { attrs: ' vMerge="1"' })), [2 * IN]));
    expect(drawn(e)).toEqual(['0,0 1x1 B']);
  });

  test('an hMerge with no span from the left and a vMerge with no span from above are drawn', async () => {
    const e = await open(
      deck(
        tr(cell('A'), cell('B', { attrs: ' hMerge="1"' }), cell('C')) +
          tr(
            cell('D', { attrs: ' vMerge="1"' }),
            cell('E'),
            cell('F', { attrs: ' hMerge="1" vMerge="1"' })
          ),
        [IN, IN, IN]
      )
    );
    expect(drawn(e)).toEqual([
      '0,0 1x1 A',
      '0,1 1x1 B',
      '0,2 1x1 C',
      '1,0 1x1 D',
      '1,1 1x1 E',
      '1,2 1x1 F',
    ]);
  });

  test('an orphan keeps its own gridSpan: the rest of a 2 x 2 merge draws as one wide cell', async () => {
    const e = await open(deck(HEADLESS, [IN, IN, IN]));
    expect(drawn(e)).toEqual(['0,0 2x1 D', '0,2 1x1 F', '1,0 1x1 G', '1,1 1x1 H', '1,2 1x1 I']);
  });

  test('a marker inside a span stays covered, whatever its kind', async () => {
    // (1,0) carries hMerge, but the rowSpan of (0,0) reaches it: it is covered.
    const e = await open(
      deck(
        tr(cell('A', { attrs: ' rowSpan="2"' }), cell('B')) +
          tr(cell('hidden', { attrs: ' hMerge="1"' }), cell('D')),
        [IN, IN]
      )
    );
    expect(drawn(e)).toEqual(['0,0 1x2 A', '0,1 1x1 B', '1,1 1x1 D']);
    expect(e.find.matches('hidden')).toHaveLength(0);
  });

  test('a span past the grid ends at the grid', async () => {
    const e = await open(
      deck(
        tr(cell('A', { attrs: ' gridSpan="5" rowSpan="9"' }), cell('', { attrs: ' hMerge="1"' })),
        [IN, IN]
      )
    );
    expect(drawn(e)).toEqual(['0,0 2x1 A']);
  });

  test('a span into an earlier block ends before it, so drawn cells never overlap', async () => {
    const e = await open(
      deck(
        tr(cell('A'), cell('B', { attrs: ' rowSpan="2"' })) +
          tr(cell('C', { attrs: ' gridSpan="2"' }), cell('', { attrs: ' vMerge="1"' })),
        [IN, IN]
      )
    );
    expect(drawn(e)).toEqual(['0,0 1x1 A', '0,1 1x2 B', '1,0 1x1 C']);
  });

  test('find reads the drawn orphan cells', async () => {
    const e = await open(deck(HEADLESS, [IN, IN, IN]));
    expect(e.find.matches('D')).toHaveLength(1);
  });

  test('the grid is the one the store reads', () => {
    const r = readPptx(deck(HEADLESS, [IN, IN, IN]));
    if (!r.ok) throw new Error('open');
    const grid = resolveTableGrid(tcsOf(xmlPartOf(r.pkg, '/ppt/slides/slide1.xml')!.root), 3);
    expect(grid.blocks.map((b) => [b.row, b.col, b.gridSpan, b.rowSpan])).toEqual([
      [0, 0, 2, 1],
      [0, 2, 1, 1],
      [1, 0, 1, 1],
      [1, 1, 1, 1],
      [1, 2, 1, 1],
    ]);
    expect([...grid.owner]).toEqual([0, 0, 1, 2, 3, 4]);
  });
});

describe('orphan cells are editable', () => {
  test('a press on an orphan cell hits it, and typing there edits its text', async () => {
    const e = await open(deck(tr(cell('A'), cell('B', { attrs: ' vMerge="1"' })), [IN, IN]));
    const frame = frameOf(e);
    const target = frame.table.cells.find((c) => c.cell.col === 1)!;
    const p = apply(frame.matrix, target.x + target.width / 2, target.y + target.height / 2);
    expect(cellAtPoint(frame, p)).toEqual({ row: 0, column: 1 });
    caretIn(e, 0, 1);
    expect(e.snapshot().selection).toMatchObject({ kind: 'text', cell: { row: 0, column: 1 } });
    expect(e.input('text', 'x').ok).toBe(true);
    expect(textOf(tcs(e)[0][1])).toBe('xB');
    // The edit keeps the file's marker: only a table command rewrites merge markup.
    expect(markup(slideRoot(e))[0]).toEqual(['', 'vMerge=1']);
  });

  test('a save with no edit, and a save after typing, keep the merge markup', async () => {
    const e = await open(deck(HEADLESS, [IN, IN, IN]));
    const before = markup(slideRoot(e));
    expect(markup(reopened(e))).toEqual(before);
    caretIn(e, 0, 0);
    expect(e.input('text', 'x').ok).toBe(true);
    const again = reopened(e);
    expect(markup(again)).toEqual(before);
    expect(textOf(tcsOf(again)[0][0])).toBe('xD');
  });

  test('row and column commands on a table with orphans leave no orphan and keep every text', async () => {
    const commands: EditorCommand[] = [
      { type: 'insertRow', where: 'below' },
      { type: 'insertColumn', where: 'right' },
      { type: 'deleteRow' },
      { type: 'deleteColumn' },
    ];
    for (const command of commands) {
      const e = await open(deck(HEADLESS, [IN, IN, IN]));
      expect(mergeViolations(e).length).toBeGreaterThan(0);
      caretIn(e, 1, 2);
      const r = e.exec(command);
      expect({ command: command.type, ok: r.ok }).toEqual({ command: command.type, ok: true });
      expect(mergeViolations(e)).toEqual([]);
      const text = textOf(find(slideRoot(e), 'a:tbl'));
      expect(text).toContain('D');
    }
  });

  test('the repair keeps what an orphan cell draws: its text and its span', async () => {
    const e = await open(deck(HEADLESS, [IN, IN, IN]));
    const shown = drawn(e);
    caretIn(e, 1, 0);
    expect(e.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    expect(drawn(e).slice(0, 2)).toEqual(shown.slice(0, 2));
    expect(markup(slideRoot(e))[0]).toEqual(['gridSpan=2', 'hMerge=1', '']);
  });

  test('merge and split with an orphan cell give a valid table', async () => {
    const e = await open(deck(HEADLESS, [IN, IN, IN]));
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 1 },
    });
    expect(e.exec({ type: 'mergeCells' }).ok).toBe(true);
    expect(mergeViolations(e)).toEqual([]);
    expect(drawn(e)[0]).toBe('0,0 2x2 DGH');
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'splitCell', rows: 2, cols: 2 }).ok).toBe(true);
    expect(mergeViolations(e)).toEqual([]);
    expect(drawn(e)).toHaveLength(6);
  });

  test('a store edit that makes a new orphan is refused; one that keeps the old orphan commits', async () => {
    const e = await open(
      deck(tr(cell('A'), cell('B', { attrs: ' vMerge="1"' }), cell('C')), [IN, IN, IN])
    );
    const store = e.session!.store;
    const set = (node: string, local: string, value: string | null): PptxOp => ({
      op: 'setAttribute',
      part: part(e),
      node,
      name: { ns: '', local },
      value,
    });
    const row = tcs(e)[0];
    // A second orphan in the same table is new: refused.
    const r = store.transact([set(row[2].id, 'hMerge', '1')]);
    expect(r).toMatchObject({
      ok: false,
      rejection: { code: 'tree-invariant', rule: 'table-merge' },
    });
    // An edit beside the old orphan commits.
    const t = find(row[0], 'a:t')!;
    expect(store.transact([{ op: 'setText', part: part(e), node: t.id, text: 'A2' }]).ok).toBe(
      true
    );
    // Removing the old orphan's marker commits too.
    expect(store.transact([set(row[1].id, 'vMerge', null)]).ok).toBe(true);
    expect(mergeViolations(e)).toEqual([]);
  });
});

describe('commands that must not write', () => {
  test('selecting a row writes nothing', async () => {
    const e = await open(deck(HEADLESS, [IN, IN, IN]));
    const revision = e.session!.store.revision;
    caretIn(e, 1, 2);
    expect(e.exec({ type: 'selectTableRegion', region: 'row' }).ok).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'cells' });
    expect(e.session!.store.revision).toBe(revision);
  });

  test('a command in a table wider than the cascade draws is refused', async () => {
    const columns = Array.from({ length: 257 }, () => IN / 100);
    const e = await open(deck(tr(...columns.map(() => cell('a'))), columns));
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(false);
  });
});

describe('copies of a table with orphans', () => {
  const tables = (root: Element): Element[] =>
    elements(find(root, 'p:spTree')!).filter((c) => c.local === 'graphicFrame');

  test('a pasted table keeps the text of every cell and drops a forged structure mark', async () => {
    const e = await open(
      deck(
        tr(
          cell('A'),
          cell('orphan', { attrs: ' vMerge="1"' }),
          cell('B', { attrs: ' gridSpan="2"' }),
          cell('covered', { attrs: ' hMerge="1"' })
        ),
        [IN, IN, IN, IN]
      )
    );
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    // A payload can carry a structure mark in a cell's extension list (a forged payload).
    const payload = e.copy()!;
    const mark: NodeSpec = {
      ns: NS.a,
      local: 'extLst',
      prefix: 'a',
      children: [
        {
          ns: NS.a,
          local: 'ext',
          prefix: 'a',
          attrs: [{ ns: '', local: 'uri', value: TABLE_EDIT_URI }],
        },
      ],
    };
    const cellsOf = (spec: NodeSpec): NodeSpec[] =>
      spec.local === 'tc'
        ? [spec]
        : (spec.children ?? []).flatMap((c) => ('text' in c ? [] : cellsOf(c)));
    const specs = cellsOf(payload.shapes[0] as NodeSpec);
    for (const i of [1, 3])
      (specs[i] as { children: NodeSpec['children'] }).children = [
        ...(specs[i].children ?? []),
        mark,
      ];
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const frames = tables(slideRoot(e));
    expect(frames).toHaveLength(2);
    const pasted = children(find(frames[1], 'a:tbl'), 'a:tr').map((r) =>
      children(r, 'a:tc').map((tc) => textOf(tc))
    );
    // Every cell keeps its text, also the covered one, as a save writes it; the marks go.
    expect(pasted).toEqual([['A', 'orphan', 'B', 'covered']]);
    expect(JSON.stringify(frames[1])).not.toContain(TABLE_EDIT_URI);
    expect(
      children(find(frames[1], 'a:tbl'), 'a:tr')
        .flatMap((r) => children(r, 'a:tc'))
        .some((tc) => children(tc, 'a:extLst').length > 0)
    ).toBe(false);
  });

  test('duplicate, paste, and duplicate slide copy the table as it draws; the source keeps its markup', async () => {
    const commands: ((e: PptxEditorInstance) => EditorCommand)[] = [
      () => ({ type: 'duplicate' }),
      (e) => ({ type: 'paste', payload: e.copy()! }),
      () => ({ type: 'duplicateSlide', index: 0 }),
    ];
    for (const make of commands) {
      const e = await open(deck(HEADLESS, [IN, IN, IN]));
      const source = markup(slideRoot(e));
      const shown = drawn(e);
      // Only the source's own violations stay.
      const old = mergeViolations(e).length;
      e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
      const command = make(e);
      expect({ command: command.type, ok: e.exec(command).ok }).toEqual({
        command: command.type,
        ok: true,
      });
      const slides = e.session!.info.slides.map(
        (s) => xmlPartOf(e.session!.store.package, s.partName)!.root
      );
      const all = slides.flatMap(tables);
      expect(all).toHaveLength(2);
      expect(markup(all[0])).toEqual(source);
      // The copy holds valid merge markup for the same drawn cells.
      expect(markup(all[1])[0]).toEqual(['gridSpan=2', 'hMerge=1', '']);
      expect(mergeViolations(e)).toHaveLength(old);
      expect(drawn(e)).toEqual(shown);
    }
  });
});
