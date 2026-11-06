// Table pointer gestures (surface-table.ts): a press in a cell places the caret, a drag across
// cells selects a cell range, Shift+press extends to a range, and a drag on a grid line of the
// table resizes its columns or rows (and selects that table). The host is a stub over a real editor and its
// layout; `TableGestures` uses no DOM.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { attr, children, elements, find, xmlPartOf, type Element } from '../../store/index.ts';
import { DEFAULT_TABLE_STYLE_ID } from '../../cascade/index.ts';
import type {
  EditorCommand,
  EditorExecOptions,
  EditorSelection,
  HistoryGroup,
} from '../../contracts/index.ts';
import { apply, type SlideLayout } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { clampTableResize, TableGestures, type TableGestureHost } from '../surface-table.ts';
import { tableItem } from '../table-cells.ts';

const IN = 914_400;
/** The narrowest column after a drag, EMU. */
const MIN_LINE = 91_440;

/** A 3 × 3 table: row 0 "A B C", row 1 "D E F", row 2 "G H I". */
function deck(columns = [IN, IN, IN], tblAttrs = ''): Uint8Array {
  const letters = (r: number) =>
    ['ABC', 'DEF', 'GHI'][r]
      .split('')
      .map((t) => cell(t))
      .join('');
  return buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, x: IN, y: 5 * IN, cx: 3 * IN, cy: IN / 2, text: para(run('Box')) }) +
          table({
            id: 3,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: 3 * 370_840,
            columns,
            tblPr: `<a:tblPr firstRow="1" bandRow="1"${tblAttrs}><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
            rows: [0, 1, 2].map((r) => `<a:tr h="370840">${letters(r)}</a:tr>`).join(''),
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
const frameId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
const boxId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'sp')!.id;
const layout = (e: PptxEditorInstance): SlideLayout => e.session!.layout(0)!;
/** The `a:gridCol` widths of the table, EMU. */
const widths = (e: PptxEditorInstance): number[] =>
  children(find(root(e), 'a:tblGrid'), 'a:gridCol').map((g) => Number(attr(g, 'w')));

const caretIn = (e: PptxEditorInstance, row: number, column: number, offset = 0): void =>
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset },
    focus: { paragraph: 0, offset },
  });

/** A host over the editor: it records each command and applies it with `exec`. */
function gestures(
  e: PptxEditorInstance,
  readOnly = false
): { g: TableGestures; commands: EditorCommand[]; groups: (HistoryGroup | undefined)[] } {
  const commands: EditorCommand[] = [];
  const groups: (HistoryGroup | undefined)[] = [];
  const host: TableGestureHost = {
    selection: () => e.snapshot().selection,
    readOnly: () => readOnly,
    zoom: () => 1,
    select: (sel: EditorSelection) => e.setSelection(sel),
    run: (command: EditorCommand, options?: EditorExecOptions) => {
      commands.push(command);
      groups.push(options?.historyGroup);
      expect(e.exec(command, options).ok).toBe(true);
    },
    beginHistoryGroup: () => e.beginHistoryGroup(),
  };
  return { g: new TableGestures(host), commands, groups };
}

/** The table frame of the current layout. */
function frameOf(e: PptxEditorInstance) {
  const frame = tableItem(layout(e), frameId(e));
  expect(frame).toBeDefined();
  return frame!;
}

/** A frame-local point in slide space. */
const at = (e: PptxEditorInstance, x: number, y: number) => apply(frameOf(e).matrix, x, y);

/** The center of a drawn cell, in slide space. */
function centerOf(e: PptxEditorInstance, row: number, column: number) {
  const c = frameOf(e).table.cells.find((x) => x.cell.row === row && x.cell.col === column)!;
  return at(e, c.x + c.width / 2, c.y + c.height / 2);
}

/** The vertical middle of a row, in frame space. */
function rowMid(e: PptxEditorInstance, row: number): number {
  const t = frameOf(e).table;
  return (t.rowY[row] + t.rowY[row + 1]) / 2;
}

describe('table gestures: caret and cell range', () => {
  test('a press inside a cell places a caret in that cell', async () => {
    const e = await open();
    const { g, commands } = gestures(e);
    expect(g.down(layout(e), frameId(e), centerOf(e, 1, 1), false)).toBe('text');
    expect(g.active).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      kind: 'text',
      nodeId: frameId(e),
      cell: { row: 1, column: 1 },
    });
    const sel = e.snapshot().selection as Extract<EditorSelection, { kind: 'text' }>;
    expect(sel.anchor).toEqual(sel.focus);
    expect(g.cursor(layout(e), frameId(e), centerOf(e, 1, 1))).toBe('text');
    g.up();
    expect(g.active).toBe(false);
    expect(commands).toEqual([]);
  });

  test('a drag across cells selects a cell range; back in the first cell it selects text', async () => {
    const e = await open();
    const { g } = gestures(e);
    const id = frameId(e);
    expect(g.down(layout(e), id, centerOf(e, 0, 0), false)).toBe('text');
    expect(g.move(layout(e), centerOf(e, 1, 2))).toBe(true);
    expect(e.snapshot().selection).toEqual({
      kind: 'cells',
      nodeId: id,
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 2 },
    });
    expect(g.move(layout(e), centerOf(e, 0, 0))).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      kind: 'text',
      nodeId: id,
      cell: { row: 0, column: 0 },
    });
    g.up();
    // No gesture runs: a move is not for the table.
    expect(g.move(layout(e), centerOf(e, 2, 2))).toBe(false);
  });

  test('Shift+press in another cell extends the caret to a cell range', async () => {
    const e = await open();
    const { g } = gestures(e);
    const id = frameId(e);
    caretIn(e, 0, 0);
    expect(g.down(layout(e), id, centerOf(e, 2, 1), true)).toBe('text');
    expect(e.snapshot().selection).toEqual({
      kind: 'cells',
      nodeId: id,
      anchor: { row: 0, column: 0 },
      focus: { row: 2, column: 1 },
    });
    g.up();
    // Shift+press with a cell range moves its focus.
    expect(g.down(layout(e), id, centerOf(e, 1, 2), true)).toBe('text');
    expect(e.snapshot().selection).toEqual({
      kind: 'cells',
      nodeId: id,
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 2 },
    });
  });

  test('a press on an element that is not a table is not a table gesture', async () => {
    const e = await open();
    const { g } = gestures(e);
    expect(g.down(layout(e), boxId(e), centerOf(e, 1, 1), false)).toBeNull();
    expect(g.active).toBe(false);
  });
});

describe('table gestures: grid lines', () => {
  test('an inner column line of the selected table resizes the two columns beside it', async () => {
    const e = await open();
    const { g, commands, groups } = gestures(e);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    const before = widths(e);
    expect(before).toEqual([IN, IN, IN]);
    const x = frameOf(e).table.columnX[1];
    const y = rowMid(e, 1);
    expect(g.cursor(layout(e), id, at(e, x, y))).toBe('col-resize');
    expect(g.down(layout(e), id, at(e, x, y), false)).toBe('grid');
    expect(g.active).toBe(true);
    // The selection stays on the table.
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [id] });

    const start = at(e, x, y);
    const moved = at(e, x + IN / 2, y);
    expect(g.move(layout(e), moved)).toBe(true);
    expect(commands).toHaveLength(1);
    const first = commands[0] as Extract<EditorCommand, { type: 'setTableGrid' }>;
    expect(first.type).toBe('setTableGrid');
    expect(first.rows).toBeUndefined();
    expect(first.columns).toEqual([IN + IN / 2, IN - IN / 2, IN]);
    expect(widths(e)).toEqual([IN + IN / 2, IN - IN / 2, IN]);
    // The table keeps its width.
    expect(widths(e).reduce((a, b) => a + b, 0)).toBe(3 * IN);

    // A huge move leaves column 1 at the narrowest width.
    expect(g.move(layout(e), { x: start.x + 20 * IN, y: start.y })).toBe(true);
    expect(commands).toHaveLength(2);
    const second = commands[1] as Extract<EditorCommand, { type: 'setTableGrid' }>;
    expect(second.columns).toEqual([2 * IN - MIN_LINE, MIN_LINE, IN]);
    // Both steps run with the one open history group of the drag; the drag's end closes it.
    expect(groups[1]).toBe(groups[0]!);
    expect(groups[0]!.state).toBe('open');
    expect(widths(e)).toEqual([2 * IN - MIN_LINE, MIN_LINE, IN]);
    g.up();
    expect(g.active).toBe(false);
    expect(groups[0]!.state).toBe('closed');

    // One drag is one undo step.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(widths(e)).toEqual(before);
  });

  test('a column line of a right-to-left table resizes the grid columns it is drawn between', async () => {
    const e = await open(deck([IN, 2 * IN, 3 * IN], ' rtl="1"'));
    const { g, commands } = gestures(e);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    // Grid column 2 (3 in) is drawn at the left, grid column 0 (1 in) at the right.
    const x = frameOf(e).table.columnX[1];
    const y = rowMid(e, 1);
    expect(g.down(layout(e), id, at(e, x, y), false)).toBe('grid');
    expect(g.move(layout(e), at(e, x + IN / 2, y))).toBe(true);
    const cmd = commands[0] as Extract<EditorCommand, { type: 'setTableGrid' }>;
    expect(cmd.columns).toEqual([IN, 2 * IN - IN / 2, 3 * IN + IN / 2]);
    expect(widths(e)).toEqual([IN, 2 * IN - IN / 2, 3 * IN + IN / 2]);
    g.up();
    caretIn(e, 0, 0);
    expect(e.snapshot().table?.columnWidth).toBe(IN);
  });

  test('the right outer edge resizes only the last column', async () => {
    const e = await open();
    const { g, commands } = gestures(e);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    const t = frameOf(e).table;
    const x = t.columnX[t.columnX.length - 1];
    const y = rowMid(e, 0);
    expect(g.cursor(layout(e), id, at(e, x, y))).toBe('col-resize');
    expect(g.down(layout(e), id, at(e, x, y), false)).toBe('grid');
    expect(g.move(layout(e), at(e, x + IN / 2, y))).toBe(true);
    const command = commands[0] as Extract<EditorCommand, { type: 'setTableGrid' }>;
    expect(command.columns).toEqual([IN, IN, IN + IN / 2]);
    expect(widths(e)).toEqual([IN, IN, IN + IN / 2]);
    g.up();
  });

  test('the right edge stops at the slide edge, where the pointer can reach it again', async () => {
    const e = await open();
    const { g } = gestures(e);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    const t = frameOf(e).table;
    const x = t.columnX[t.columnX.length - 1];
    const y = rowMid(e, 0);
    expect(g.down(layout(e), id, at(e, x, y), false)).toBe('grid');
    // Far past the slide: the table at 1 in, 3 in wide, grows to the slide's right edge only.
    expect(g.move(layout(e), at(e, x + 20 * IN, y))).toBe(true);
    const slide = layout(e).slide.width;
    expect(IN + widths(e).reduce((a, b) => a + b, 0)).toBe(slide);
    g.up();
  });

  test('a frame resize handle stops at the slide edge too; other shapes and handles do not', async () => {
    const e = await open();
    const frame = frameOf(e);
    const slide = layout(e).slide;
    const right = frame.matrix[4] + frame.element.xfrm.cx;
    const bottom = frame.matrix[5] + frame.element.xfrm.cy;
    // The east and south-east handles grow the table to the slide's edge only.
    expect(clampTableResize(frame, [1, 0.5], 20 * IN, 0, slide)).toEqual({
      dx: slide.width - right,
      dy: 0,
    });
    expect(clampTableResize(frame, [1, 1], 20 * IN, 20 * IN, slide)).toEqual({
      dx: slide.width - right,
      dy: slide.height - bottom,
    });
    // A shrink, the west handle, and a shape that is not a table keep the pointer's move.
    expect(clampTableResize(frame, [1, 0.5], -IN, 0, slide)).toEqual({ dx: -IN, dy: 0 });
    expect(clampTableResize(frame, [0, 0.5], -20 * IN, 0, slide)).toEqual({ dx: -20 * IN, dy: 0 });
    // Shift on a corner keeps the aspect: the scale stops where the first edge meets the slide.
    const { cx, cy } = frame.element.xfrm;
    const grow = Math.min((slide.width - right) / cx, (slide.height - bottom) / cy);
    const kept = clampTableResize(frame, [1, 1], 0, 20 * IN, slide, true);
    expect(kept.dx).toBe(0);
    expect(kept.dy).toBeCloseTo(grow * cy);
    expect(Math.max(1 + kept.dx / cx, 1 + kept.dy / cy) * cx).toBeLessThanOrEqual(
      slide.width - frame.matrix[4] + 1
    );
    const box = layout(e).items.find((i) => i.element.nodeId === boxId(e))!;
    expect(clampTableResize(box, [1, 1], 20 * IN, 20 * IN, slide)).toEqual({
      dx: 20 * IN,
      dy: 20 * IN,
    });
  });

  test('the outer edges are grid lines, so the border band of the selected table yields', async () => {
    const e = await open();
    const { g } = gestures(e);
    const frame = frameOf(e);
    const t = frame.table;
    const y = rowMid(e, 0);
    expect(g.onLine(frame, at(e, t.columnX.at(-1)!, y))).toBe(true);
    expect(g.onLine(frame, at(e, (t.columnX[0] + t.columnX[1]) / 2, y))).toBe(false);
    // A shape that is not a table never has grid lines.
    const box = layout(e).items.find((i) => i.element.nodeId === boxId(e))!;
    expect(g.onLine(box, at(e, t.columnX.at(-1)!, y))).toBe(false);
    expect(gestures(e, true).g.onLine(frame, at(e, t.columnX.at(-1)!, y))).toBe(false);
    // Just outside the right edge, the hit test finds the selected table only.
    const outside = at(e, t.columnX.at(-1)! + 2 * 9_525, y);
    expect(g.edgeAt(layout(e), outside)).toBeUndefined();
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    expect(g.edgeAt(layout(e), outside)).toBe(frameId(e));
    expect(g.edgeAt(layout(e), at(e, -2 * 9_525, y))).toBeUndefined();
  });

  test('a row line shows the row cursor and a drag resizes the rows', async () => {
    const e = await open();
    const { g, commands } = gestures(e);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    const t = frameOf(e).table;
    const heights = t.rowY.slice(1).map((v, i) => Math.round(v - t.rowY[i]));
    const x = (t.columnX[0] + t.columnX[1]) / 2;
    const y = t.rowY[1];
    expect(g.cursor(layout(e), id, at(e, x, y))).toBe('row-resize');
    expect(g.down(layout(e), id, at(e, x, y), false)).toBe('grid');
    expect(g.move(layout(e), at(e, x, y + IN / 4))).toBe(true);
    const command = commands[0] as Extract<EditorCommand, { type: 'setTableGrid' }>;
    expect(command.columns).toBeUndefined();
    expect(command.rows).toEqual([heights[0] + IN / 4, heights[1], heights[2]]);
    g.up();
  });

  test('a grid line of a table that is not selected drags too, as the chrome marks it', async () => {
    const e = await open();
    const { g } = gestures(e);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [boxId(e)] });
    const p = at(e, frameOf(e).table.columnX[1], rowMid(e, 1));
    expect(g.cursor(layout(e), id, p)).toBe('col-resize');
    expect(g.down(layout(e), id, p, false)).toBe('grid');
    // A press alone selects nothing; the first move selects the table, and the drag resizes.
    expect(e.snapshot().selection).toMatchObject({ kind: 'shapes', nodeIds: [boxId(e)] });
    g.move(layout(e), at(e, frameOf(e).table.columnX[1] + IN / 4, rowMid(e, 1)));
    expect(e.snapshot().selection).toMatchObject({ kind: 'shapes', nodeIds: [id] });
    g.up();
    expect(widths(e)).toEqual([IN + IN / 4, IN - IN / 4, IN]);
  });

  test('a click on a line of a table that is not selected changes nothing', async () => {
    const e = await open();
    const { g, commands } = gestures(e);
    e.setSelection({ kind: 'shapes', nodeIds: [boxId(e)] });
    const p = at(e, frameOf(e).table.columnX[1], rowMid(e, 1));
    expect(g.down(layout(e), frameId(e), p, false)).toBe('grid');
    g.up();
    expect(commands).toEqual([]);
    expect(e.snapshot().selection).toMatchObject({ kind: 'shapes', nodeIds: [boxId(e)] });
  });

  test('read-only: a press is not a table gesture and shows no cursor', async () => {
    const e = await open();
    const { g, commands } = gestures(e, true);
    const id = frameId(e);
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    const line = at(e, frameOf(e).table.columnX[1], rowMid(e, 1));
    expect(g.down(layout(e), id, line, false)).toBeNull();
    expect(g.down(layout(e), id, centerOf(e, 1, 1), false)).toBeNull();
    expect(g.cursor(layout(e), id, line)).toBeUndefined();
    expect(g.active).toBe(false);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [id] });
    expect(commands).toEqual([]);
  });
});
