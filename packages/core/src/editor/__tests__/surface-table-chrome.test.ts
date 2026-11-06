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
import { tableItem } from '../table-cells.ts';

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
    // The caret went into row 0, so the new row is row 1.
    expect(e.snapshot().selection).toMatchObject({ nodeId: frameId(e), cell: { row: 0 } });
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
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 0, column: 0 } });
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

  test('a selected table offers no column button under its rotation handle', async () => {
    const e = await open(deck(false, 4));
    const { c } = chrome(e);
    // Hovered only: the middle line has its button.
    c.hover(layout(e), at(e, columnX(e)[2], -10 * PX), frameId(e));
    expect(classes(c.elements(doc, layout(e)))).toContain('pptx-surface__table-insert');
    // Selected: the rotation handle sits over the top center, so the button stays away.
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    c.hover(layout(e), at(e, columnX(e)[2], -10 * PX), frameId(e));
    expect(classes(c.elements(doc, layout(e)))).toEqual(['pptx-surface__table-handle']);
    c.hover(layout(e), at(e, columnX(e)[1], -10 * PX), frameId(e));
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
