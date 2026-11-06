// The table chrome (surface-table-chrome.ts): the select handle at the table's top-left corner,
// the mark on the grid line under the pointer, and the "+" buttons that insert a row or a column.
// The host is a stub over a real editor and its layout; the overlay elements go into happy-dom.

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  cell,
  shape,
  para,
  run,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { children, elements, find, xmlPartOf, type Element } from '../../store/index.ts';
import { DEFAULT_TABLE_STYLE_ID } from '../../cascade/index.ts';
import type { EditorCommand, EditorSelection } from '../../contracts/index.ts';
import { apply, type SlideLayout } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { TableChrome, type TableChromeHost } from '../surface-table-chrome.ts';
import { gridLineAt, tableItem } from '../table-cells.ts';
import { cellsOverlay } from '../surface-table.ts';

const IN = 914_400;
const ROW = 370_840;
/** One CSS pixel at zoom 1, EMU. */
const PX = 9_525;

/** A 3 × 3 table at (1", 1"), and a text box below it. `merge` spans row 1's first two cells. */
function deck(merge = false, columns = 3): Uint8Array {
  const rows = [0, 1, 2].map((r) => {
    const cells =
      merge && r === 1
        ? '<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>' +
          cell('F')
        : ['A', 'B', 'C', 'D']
            .slice(0, columns)
            .map((t) => cell(t))
            .join('');
    return `<a:tr h="${ROW}">${cells}</a:tr>`;
  });
  return buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, x: IN, y: 5 * IN, cx: 3 * IN, cy: IN / 2, text: para(run('Box')) }) +
          table({
            id: 3,
            x: IN,
            y: IN,
            cx: columns * IN,
            cy: 3 * ROW,
            columns: Array.from({ length: columns }, () => IN),
            tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
            rows: rows.join(''),
          }),
      },
    ],
  });
}

async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const part = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const root = (e: PptxEditorInstance): Element => xmlPartOf(e.session!.store.package, part(e))!.root;
const spTree = (e: PptxEditorInstance) => elements(find(root(e), 'p:spTree')!);
const frameId = (e: PptxEditorInstance) => spTree(e).find((c) => c.local === 'graphicFrame')!.id;
const boxId = (e: PptxEditorInstance) => spTree(e).find((c) => c.local === 'sp')!.id;
const layout = (e: PptxEditorInstance): SlideLayout => e.session!.layout(0)!;
const rowCount = (e: PptxEditorInstance) => children(find(root(e), 'a:tbl')!, 'a:tr').length;
const colCount = (e: PptxEditorInstance) =>
  children(find(root(e), 'a:tblGrid')!, 'a:gridCol').length;
const doc = new GlobalWindow().document as unknown as Document;

function chrome(e: PptxEditorInstance, readOnly = false, painting = { on: false }) {
  const commands: EditorCommand[] = [];
  const host: TableChromeHost = {
    selection: () => e.snapshot().selection,
    readOnly: () => readOnly,
    painting: () => painting.on,
    zoom: () => 1,
    label: (key) => key,
    select: (sel: EditorSelection) => e.setSelection(sel),
    run: (command) => {
      commands.push(command);
      expect(e.exec(command).ok).toBe(true);
    },
  };
  return { c: new TableChrome(host), commands };
}

/** A slide point at a table-local point. */
function at(e: PptxEditorInstance, x: number, y: number): { x: number; y: number } {
  return apply(tableItem(layout(e), frameId(e))!.matrix, x, y);
}

const rowY = (e: PptxEditorInstance) => tableItem(layout(e), frameId(e))!.table.rowY;
const columnX = (e: PptxEditorInstance) => tableItem(layout(e), frameId(e))!.table.columnX;
const classes = (parts: globalThis.Element[]) => parts.map((p) => p.getAttribute('class')).sort();

describe('the table chrome', () => {
  test('shows nothing until the pointer is over a table or the selection is in one', async () => {
    const e = await open();
    const { c } = chrome(e);
    expect(c.elements(doc, layout(e))).toEqual([]);
    e.setSelection({ kind: 'shapes', nodeIds: [boxId(e)] });
    expect(c.elements(doc, layout(e))).toEqual([]);
    // Hovering a cell shows the select handle.
    expect(c.hover(layout(e), at(e, IN / 2, rowY(e)[1] / 2), frameId(e))).toBe(true);
    expect(classes(c.elements(doc, layout(e)))).toEqual(['pptx-surface__table-handle']);
    // The same place again changes nothing; far away the hovered table's chrome goes.
    expect(c.hover(layout(e), at(e, IN / 2, rowY(e)[1] / 2), frameId(e))).toBe(false);
    expect(c.hover(layout(e), at(e, 6 * IN, 6 * IN), undefined)).toBe(true);
    expect(c.elements(doc, layout(e))).toEqual([]);
  });

  test('a caret in a cell keeps the handle at the top-left corner, 22 px out', async () => {
    const e = await open();
    const { c } = chrome(e);
    const point = { paragraph: 0, offset: 0 };
    e.setSelection({
      kind: 'text',
      nodeId: frameId(e),
      cell: { row: 1, column: 1 },
      anchor: point,
      focus: point,
    });
    const [handle] = c.elements(doc, layout(e));
    expect(handle.getAttribute('data-table-handle')).toBe(frameId(e));
    const corner = at(e, 0, 0);
    expect(handle.getAttribute('transform')).toBe(
      `translate(${corner.x - 22 * PX} ${corner.y - 22 * PX})`
    );
    expect(handle.querySelector('title')?.textContent).toBe('table.region.table');
  });

  test('a press on the handle asks the surface to select and move the table', async () => {
    const e = await open();
    const { c, commands } = chrome(e);
    c.hover(layout(e), at(e, IN / 2, IN / 10), frameId(e));
    const [handle] = c.elements(doc, layout(e));
    const box = handle.querySelector('rect')!;
    expect(c.press(layout(e), box)).toEqual({ kind: 'move', id: frameId(e), scope: {} });
    expect(commands).toEqual([]);
    // Anything else is not the chrome.
    expect(c.press(layout(e), doc.createElement('div') as unknown as globalThis.Element)).toBe(
      null
    );
  });

  test('the grid line under the pointer is marked across the table', async () => {
    const e = await open();
    const { c } = chrome(e);
    c.hover(layout(e), at(e, columnX(e)[1] + 2 * PX, rowY(e)[1] + ROW / 2), frameId(e));
    const parts = c.elements(doc, layout(e));
    expect(classes(parts)).toEqual(['pptx-surface__table-handle', 'pptx-surface__table-line']);
    const line = parts.find((p) => p.getAttribute('class') === 'pptx-surface__table-line')!;
    const top = at(e, columnX(e)[1], 0);
    const bottom = at(e, columnX(e)[1], rowY(e).at(-1)!);
    expect(Number(line.getAttribute('x1'))).toBeCloseTo(top.x);
    expect(Number(line.getAttribute('y1'))).toBeCloseTo(top.y);
    expect(Number(line.getAttribute('y2'))).toBeCloseTo(bottom.y);
  });

  test('the left edge offers a row below the nearest line; a press inserts it there', async () => {
    const e = await open();
    const { c, commands } = chrome(e);
    // Outside the table, 10 px left of its edge, near the line under row 0.
    c.hover(layout(e), at(e, -10 * PX, rowY(e)[1] - 3 * PX), undefined);
    expect(classes(c.elements(doc, layout(e)))).toEqual([]);
    // The selected table keeps its chrome outside its box, so its buttons can be reached.
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    c.hover(layout(e), at(e, -10 * PX, rowY(e)[1] - 3 * PX), undefined);
    const parts = c.elements(doc, layout(e));
    expect(classes(parts)).toEqual([
      'pptx-surface__table-handle',
      'pptx-surface__table-insert',
      'pptx-surface__table-insert-guide',
    ]);
    const button = parts.find((p) => p.getAttribute('class') === 'pptx-surface__table-insert')!;
    expect(button.getAttribute('data-table-insert')).toBe(`${frameId(e)}|row|1`);
    expect(button.querySelector('title')?.textContent).toBe('table.insertRowBelow');
    expect(c.press(layout(e), button.querySelector('circle')!)).toEqual({ kind: 'insert' });
    expect(commands).toEqual([{ type: 'insertRow', where: 'below' }]);
    expect(rowCount(e)).toBe(4);
    // The new row is row 1, and the caret is in it.
    expect(e.snapshot().selection).toMatchObject({ nodeId: frameId(e), cell: { row: 1 } });
  });

  test('the top edge offers a column right of the nearest line', async () => {
    const e = await open();
    const { c, commands } = chrome(e);
    c.hover(layout(e), at(e, columnX(e)[2] + 5 * PX, 3 * PX), frameId(e));
    const button = c
      .elements(doc, layout(e))
      .find((p) => p.getAttribute('class') === 'pptx-surface__table-insert')!;
    expect(button.getAttribute('data-table-insert')).toBe(`${frameId(e)}|column|2`);
    c.press(layout(e), button);
    expect(commands).toEqual([{ type: 'insertColumn', where: 'right' }]);
    expect(colCount(e)).toBe(4);
    // The caret is in the new column, in the first row.
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 0, column: 2 } });
  });

  test('a refused insert leaves the selection as it was', async () => {
    const e = await open();
    const selections: EditorSelection[] = [];
    const c = new TableChrome({
      selection: () => e.snapshot().selection,
      readOnly: () => false,
      zoom: () => 1,
      label: (key) => key,
      select: (sel: EditorSelection) => (selections.push(sel), e.setSelection(sel)),
      run: () => ({ ok: false }),
    });
    const before: EditorSelection = { kind: 'shapes', nodeIds: [frameId(e)] };
    e.setSelection(before);
    c.hover(layout(e), at(e, columnX(e)[2] + 5 * PX, 3 * PX), frameId(e));
    const button = c
      .elements(doc, layout(e))
      .find((p) => p.getAttribute('class') === 'pptx-surface__table-insert')!;
    expect(c.press(layout(e), button)).toEqual({ kind: 'insert' });
    expect(selections.at(-1)).toEqual(before);
    expect(e.snapshot().selection).toMatchObject(before);
  });

  test('a selected column paints its own cells, not the merged cell that crosses it', async () => {
    const e = await open(deck(true));
    const nodeId = frameId(e);
    const rects = (sel: EditorSelection) =>
      cellsOverlay(doc, layout(e), sel)[0].querySelectorAll('.pptx-surface__cell-selection').length;
    const edges = (sel: EditorSelection) =>
      cellsOverlay(doc, layout(e), sel)[0].querySelectorAll('.pptx-surface__cell-selection-edge');
    // Column 1: rows 0 and 2 have their own cell; row 1's merged cell starts in column 0.
    const column = {
      kind: 'cells',
      nodeId,
      anchor: { row: 0, column: 1 },
      focus: { row: 2, column: 1 },
    } as const;
    expect(rects(column)).toBe(2);
    // A range that is not a whole row or column widens over the merged cell.
    const block = {
      kind: 'cells',
      nodeId,
      anchor: { row: 0, column: 1 },
      focus: { row: 1, column: 1 },
    } as const;
    expect(rects(block)).toBe(3);
    // One edge around the whole range, not one on each cell.
    expect(edges(block).length).toBe(1);
    // Column 0, where the merged cell starts: the edge is one column wide, as the command's.
    const first = { ...column, anchor: { row: 0, column: 0 }, focus: { row: 2, column: 0 } };
    const t = tableItem(layout(e), nodeId)!.table;
    const edge = edges(first)[0];
    expect(Number(edge.getAttribute('width'))).toBeCloseTo(t.columnX[1] - t.columnX[0]);
    expect(Number(edge.getAttribute('height'))).toBeCloseTo(t.rowY[3] - t.rowY[0]);
    // The merged cell that starts in column 0 is tinted inside the column only, within the edge.
    const tints = [
      ...cellsOverlay(doc, layout(e), first)[0].querySelectorAll('.pptx-surface__cell-selection'),
    ];
    for (const r of tints)
      expect(Number(r.getAttribute('width'))).toBeCloseTo(t.columnX[1] - t.columnX[0]);
  });

  test('a selected column under a merged cell inserts, deletes, and fills its own line', async () => {
    const select = async () => {
      const e = await open(deck(true));
      const point = { paragraph: 0, offset: 0 };
      e.setSelection({
        kind: 'text',
        nodeId: frameId(e),
        cell: { row: 0, column: 1 },
        anchor: point,
        focus: point,
      });
      expect(e.exec({ type: 'selectTableRegion', region: 'column' }).ok).toBe(true);
      return e;
    };
    for (const where of ['left', 'right'] as const) {
      const e = await select();
      expect(e.exec({ type: 'insertColumn', where }).ok).toBe(true);
      expect(colCount(e)).toBe(4);
    }
    const e = await select();
    expect(e.exec({ type: 'deleteColumn' }).ok).toBe(true);
    expect(colCount(e)).toBe(2);
    // Fill changes the two cells the highlight shows, not row 1's merged cell.
    const f = await select();
    expect(f.exec({ type: 'setCellFill', color: '#FF0000' }).ok).toBe(true);
    const filled = find(root(f), 'a:tbl')!;
    expect(
      children(filled, 'a:tr').map((tr) =>
        children(tr, 'a:tc').map((tc) => !!find(tc, 'a:srgbClr'))
      )
    ).toEqual([
      [false, true, false],
      [false, false, false],
      [false, true, false],
    ]);
  });

  test('a grid line hits only where it is drawn, not inside a merged cell', async () => {
    const e = await open(deck(true));
    const frame = tableItem(layout(e), frameId(e))!;
    const line = columnX(e)[1];
    // Row 0: the line between columns 0 and 1 is drawn there.
    expect(gridLineAt(frame, at(e, line, rowY(e)[0] + ROW / 2), 2 * PX)).toEqual({
      axis: 'column',
      index: 1,
    });
    // Row 1: its merged cell spans columns 0 and 1, so the line is not there.
    expect(gridLineAt(frame, at(e, line, rowY(e)[1] + ROW / 2), 2 * PX)).toBeUndefined();
  });

  test('a line inside a merged cell inserts from the cell that ends there', async () => {
    const e = await open(deck(true));
    const { c, commands } = chrome(e);
    // Column line 1 runs through row 1's merged cell; row 0's cell A ends at it.
    c.hover(layout(e), at(e, columnX(e)[1], -10 * PX), frameId(e));
    const button = c
      .elements(doc, layout(e))
      .find((p) => p.getAttribute('class') === 'pptx-surface__table-insert')!;
    c.press(layout(e), button);
    expect(commands).toEqual([{ type: 'insertColumn', where: 'right' }]);
    // The insert went from row 0's cell A; the caret is in the new column, in row 0.
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 0, column: 1 } });
    expect(colCount(e)).toBe(4);
  });

  test('the pointer keeps the chrome of a hovered table on its way to the handle', async () => {
    const e = await open();
    const { c } = chrome(e);
    e.setSelection({ kind: 'shapes', nodeIds: [boxId(e)] });
    c.hover(layout(e), at(e, IN / 2, IN / 10), frameId(e));
    // Out from the corner, over no element: the handle's box is 6 to 22 px out.
    for (const out of [3, 10, 14, 21]) {
      c.hover(layout(e), at(e, -out * PX, -out * PX), undefined);
      expect(classes(c.elements(doc, layout(e)))).toEqual(['pptx-surface__table-handle']);
    }
    // Past the handle, it goes.
    c.hover(layout(e), at(e, -40 * PX, -40 * PX), undefined);
    expect(c.elements(doc, layout(e))).toEqual([]);
  });

  test('a selected table offers the column button on its middle line', async () => {
    const e = await open(deck(false, 4));
    const { c } = chrome(e);
    // A table has no rotation handle, so the top center keeps its button when selected.
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    c.hover(layout(e), at(e, columnX(e)[2], -10 * PX), frameId(e));
    expect(classes(c.elements(doc, layout(e)))).toContain('pptx-surface__table-insert');
  });

  test('a line a remote edit removed is not drawn', async () => {
    const e = await open();
    const { c } = chrome(e);
    c.hover(layout(e), at(e, columnX(e)[2] + 2 * PX, rowY(e)[1] + ROW / 2), frameId(e));
    expect(classes(c.elements(doc, layout(e)))).toContain('pptx-surface__table-line');
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 1 },
      focus: { row: 2, column: 2 },
    });
    expect(e.exec({ type: 'deleteColumn' }).ok).toBe(true);
    expect(classes(c.elements(doc, layout(e)))).toEqual(['pptx-surface__table-handle']);
  });

  test('the format painter hides the chrome', async () => {
    const e = await open();
    const painting = { on: true };
    const { c } = chrome(e, false, painting);
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    expect(c.hover(layout(e), at(e, IN / 2, IN / 10), frameId(e))).toBe(false);
    expect(c.elements(doc, layout(e))).toEqual([]);
    painting.on = false;
    expect(classes(c.elements(doc, layout(e)))).toEqual(['pptx-surface__table-handle']);
  });

  test('read-only: no chrome, and presses are not for it', async () => {
    const e = await open();
    const { c } = chrome(e, true);
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    expect(c.hover(layout(e), at(e, IN / 2, IN / 10), frameId(e))).toBe(false);
    expect(c.elements(doc, layout(e))).toEqual([]);
  });
});
