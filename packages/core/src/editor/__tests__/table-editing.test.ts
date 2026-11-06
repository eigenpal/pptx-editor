// Table editing (`parity-baseline` 4.1 and 4.2, `editor-ui-parity` G6): text in cells with
// the existing text commands, Tab and Shift+Tab between cells, insert a table, insert and delete
// rows and columns, merge and split, and table and cell formatting. Every structure edit is
// checked for a valid table (one cell per grid column, merge markers as 21.1.3.16 tc defines)
// and survives save and reopen.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  children,
  elements,
  find,
  readPptx,
  semanticDigest,
  canonicalPackageFingerprint,
  textOf,
  writePptx,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { DEFAULT_TABLE_STYLE_ID, resolveSlide } from '../../cascade/index.ts';
import type { EditorSelection } from '../../contracts/index.ts';
import {
  createPptxEditor,
  runToolbarCommand,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { applyTableKey, tableKeyAction } from '../table-keys.ts';
import {
  blockAttrs,
  blockOf,
  readTable,
  type TableModel,
} from '../../automation/commands/table-model.ts';

const IN = 914_400;

/** A 3 × 3 table: row 0 "A B C", row 1 "D E F", row 2 "G H I". */
function deck(rows?: string, columns = [IN, IN, IN]): Uint8Array {
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
            cx: columns.reduce((a, b) => a + b, 0),
            cy: 3 * 370_840,
            columns,
            tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
            rows: rows ?? [0, 1, 2].map((r) => `<a:tr h="370840">${letters(r)}</a:tr>`).join(''),
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
const xml = (e: PptxEditorInstance): string =>
  JSON.stringify(xmlPartOf(e.session!.store.package, part(e))!.root);
const root = (e: PptxEditorInstance): Element => xmlPartOf(e.session!.store.package, part(e))!.root;
const frameId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
const model = (e: PptxEditorInstance): TableModel =>
  readTable(new OpBatch(e.session!.store.package), part(e), frameId(e));
/** The text of every cell, row by row (covered cells too). */
const grid = (e: PptxEditorInstance): string[][] =>
  model(e).tcs.map((row) => row.map((tc) => textOf(child(tc, 'a:txBody'))));

/**
 * The table is valid: one `a:tc` per grid column in every row (readTable checks), every cell
 * carries exactly the merge attributes of its block, and the cascade draws no orphan.
 */
function expectValid(e: PptxEditorInstance): void {
  const m = model(e);
  for (let r = 0; r < m.nR; r += 1)
    for (let c = 0; c < m.nC; c += 1) {
      const want = blockAttrs(blockOf(m, { row: r, column: c }), r, c);
      const tc = m.tcs[r][c];
      const have: Record<string, string> = {};
      for (const name of ['gridSpan', 'rowSpan', 'hMerge', 'vMerge']) {
        const v = attr(tc, name);
        if (v !== undefined) have[name] = v;
      }
      expect({ r, c, ...have }).toEqual({ r, c, ...want });
    }
  const resolved = resolveSlide(e.session!.store.package, 0)!;
  const frame = resolved.elements.find((el) => el.kind === 'graphicFrame');
  const drawn = frame?.kind === 'graphicFrame' ? frame.table!.cells.filter((x) => !x.covered) : [];
  const anchors = new Set(m.anchors.flat().map((a) => `${a.row}:${a.column}`));
  expect(drawn.map((x) => `${x.row}:${x.col}`).sort()).toEqual([...anchors].sort());
}

/** Save, reopen, and compare the semantic digest. */
function roundTrip(e: PptxEditorInstance): void {
  const bytes = e.save()!;
  const r = readPptx(bytes);
  expect(r.ok).toBe(true);
  if (!r.ok) return;
  expect(semanticDigest(r.pkg)).toBe(semanticDigest(e.session!.store.package));
  const again = readPptx(writePptx(r.pkg));
  expect(again.ok && semanticDigest(again.pkg)).toBe(semanticDigest(r.pkg));
}

const caretIn = (e: PptxEditorInstance, row: number, column: number, offset = 0): void =>
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset },
    focus: { paragraph: 0, offset },
  });

function press(e: PptxEditorInstance, key: string, shift = false): boolean {
  const action = tableKeyAction(e.session!.layout(0), e.snapshot().selection, key, shift, false);
  if (!action) return false;
  applyTableKey(action, e.exec, () => e.session!.layout(0), e.setSelection);
  return true;
}

describe('cell text (parity-baseline 4.1)', () => {
  test('typing in a cell edits that cell and undoes as one step; save and reopen keep it', async () => {
    const e = await open();
    caretIn(e, 1, 1, 1);
    for (const ch of 'xy') expect(e.input('text', ch).ok).toBe(true);
    expect(grid(e)[1]).toEqual(['D', 'Exy', 'F']);
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 1, column: 1 },
      focus: { offset: 3 },
    });
    e.input('paragraph');
    e.input('text', 'z');
    const body = child(model(e).tcs[1][1], 'a:txBody')!;
    expect(children(body, 'a:p').map((p) => textOf(p))).toEqual(['Exy', 'z']);
    roundTrip(e);
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    expect(grid(e)[1]).toEqual(['D', 'E', 'F']);
  });

  test('text commands format the cell range, the cell, or every cell of a selected table', async () => {
    const e = await open();
    // Row 0 is the header row: the style makes it bold, so Bold turns it off there.
    caretIn(e, 0, 0);
    expect(e.snapshot().format?.bold).toBe(true);
    caretIn(e, 1, 0);
    const sel = e.snapshot().selection as Extract<EditorSelection, { kind: 'text' }>;
    e.setSelection({ ...sel, focus: { paragraph: 0, offset: 1 } });
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    const bold = (r: number, c: number) =>
      attr(find(child(model(e).tcs[r][c], 'a:txBody'), 'a:rPr'), 'b') === '1';
    expect([bold(1, 0), bold(1, 1)]).toEqual([true, false]);
    expect(e.snapshot().format?.bold).toBe(true);
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 1, column: 0 },
      focus: { row: 1, column: 1 },
    });
    expect(e.exec({ type: 'setFontSize', points: 28 }).ok).toBe(true);
    const size = (r: number, c: number) =>
      attr(find(child(model(e).tcs[r][c], 'a:txBody'), 'a:rPr'), 'sz');
    expect([size(1, 0), size(1, 1), size(1, 2)]).toEqual(['2800', '2800', undefined]);
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    expect(e.exec({ type: 'toggleItalic' }).ok).toBe(true);
    const italic = model(e)
      .tcs.flat()
      .map((tc) => attr(find(tc, 'a:rPr'), 'i'));
    expect(italic.every((v) => v === '1')).toBe(true);
    roundTrip(e);
  });

  test('a row grows to fit the cell text, and the frame follows the table', async () => {
    const e = await open();
    const cy = () => Number(attr(find(find(root(e), 'p:graphicFrame')!, 'a:ext'), 'cy'));
    const before = cy();
    caretIn(e, 0, 0, 1);
    for (let i = 0; i < 3; i += 1) e.input('paragraph');
    expect(cy()).toBeGreaterThan(before);
    // The row height stays the minimum.
    expect(attr(model(e).rows[0], 'h')).toBe('370840');
  });

  test('Tab moves to the next cell and selects its text; Shift+Tab goes back', async () => {
    const e = await open();
    caretIn(e, 0, 2, 1);
    expect(press(e, 'Tab')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 1, column: 0 },
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 1 },
    });
    press(e, 'Tab', true);
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 0, column: 2 } });
    // Typing replaces the selected cell text.
    press(e, 'Tab');
    e.input('text', 'Q');
    expect(grid(e)[1][0]).toBe('Q');
  });

  test('Tab in the last cell adds a row and moves into it; undo removes the row', async () => {
    const e = await open();
    caretIn(e, 2, 2, 1);
    press(e, 'Tab');
    expect(model(e).nR).toBe(4);
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 3, column: 0 } });
    e.input('text', 'J');
    expect(grid(e)[3]).toEqual(['J', '', '']);
    expectValid(e);
    roundTrip(e);
    e.exec({ type: 'undo' });
    e.exec({ type: 'undo' });
    expect(model(e).nR).toBe(3);
  });

  test('arrow keys at the edge of a cell move into the next cell', async () => {
    const e = await open();
    caretIn(e, 0, 0, 1);
    expect(press(e, 'ArrowRight')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 0, column: 1 },
      focus: { offset: 0 },
    });
    expect(press(e, 'ArrowLeft')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 0, column: 0 },
      focus: { offset: 1 },
    });
    expect(press(e, 'ArrowDown')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 1, column: 0 } });
    // Inside the text, arrows stay with text editing.
    expect(press(e, 'ArrowLeft')).toBe(false);
  });

  test('Escape in a cell selects the table; Delete with cells selected clears their text', async () => {
    const e = await open();
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    press(e, 'Delete');
    expect(grid(e)[0]).toEqual(['', '', 'C']);
    expect(model(e).nR).toBe(3);
    press(e, 'Escape');
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [frameId(e)] });
  });
});

describe('table structure (parity-baseline 4.2)', () => {
  test('insert a table: the default size, the default style, a caret in the first cell', async () => {
    const e = await open();
    expect(e.exec({ type: 'insertTable', rows: 2, cols: 4 }).ok).toBe(true);
    const frames = elements(find(root(e), 'p:spTree')!).filter((c) => c.local === 'graphicFrame');
    const t = find(frames[1], 'a:tbl')!;
    expect(children(find(t, 'a:tblGrid'), 'a:gridCol').map((g) => attr(g, 'w'))).toEqual(
      Array(4).fill(String(Math.floor(8_128_000 / 4)))
    );
    expect(children(t, 'a:tr').map((tr) => attr(tr, 'h'))).toEqual(['370840', '370840']);
    expect(textOf(find(t, 'a:tableStyleId'))).toBe(DEFAULT_TABLE_STYLE_ID);
    expect(attr(find(t, 'a:tblPr'), 'firstRow')).toBe('1');
    expect(e.snapshot().selection).toMatchObject({
      kind: 'text',
      nodeId: frames[1].id,
      cell: { row: 0, column: 0 },
    });
    e.input('text', 'Hi');
    expect(textOf(find(find(t, 'a:tr'), 'a:txBody'))).toBe('');
    roundTrip(e);
    expect(e.can({ type: 'insertTable', rows: 0, cols: 2 }).ok).toBe(false);
    expect(e.can({ type: 'insertTable', rows: 2, cols: 76 }).ok).toBe(false);
  });

  test('insert rows and columns on every side; the caret stays in its cell', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    e.exec({ type: 'insertRow', where: 'above' });
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 2, column: 1 } });
    e.exec({ type: 'insertRow', where: 'below' });
    e.exec({ type: 'insertColumn', where: 'left' });
    expect(e.snapshot().selection).toMatchObject({ cell: { row: 2, column: 2 } });
    e.exec({ type: 'insertColumn', where: 'right' });
    expect(grid(e).map((r) => r.join(''))).toEqual(['ABC', '', 'DEF', '', 'GHI']);
    expect(model(e).nC).toBe(5);
    // The table keeps its width of three inches; the five columns share it.
    const widths = model(e).gridCols.map((g) => Number(attr(g, 'w')));
    expect(widths.reduce((a, b) => a + b, 0)).toBe(3 * IN);
    expect(Math.max(...widths) - Math.min(...widths)).toBeLessThanOrEqual(1);
    expect(e.snapshot().table).toMatchObject({
      rows: 5,
      columns: 5,
      selectedRows: { from: 2, to: 2 },
      selectedColumns: { from: 2, to: 2 },
    });
    expectValid(e);
    roundTrip(e);
  });

  test('delete rows and columns; deleting all rows deletes the table', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    e.exec({ type: 'deleteRow' });
    expect(grid(e).map((r) => r.join(''))).toEqual(['ABC', 'GHI']);
    e.exec({ type: 'deleteColumn' });
    expect(grid(e).map((r) => r.join(''))).toEqual(['AC', 'GI']);
    expectValid(e);
    roundTrip(e);
    e.exec({ type: 'selectTableRegion', region: 'column' });
    expect(e.snapshot().selection).toMatchObject({
      kind: 'cells',
      anchor: { row: 0 },
      focus: { row: 1 },
    });
    e.exec({ type: 'selectTableRegion', region: 'table' });
    e.exec({ type: 'deleteRow' });
    expect(elements(find(root(e), 'p:spTree')!).some((c) => c.local === 'graphicFrame')).toBe(
      false
    );
    expect(e.snapshot().selection).toEqual({ kind: 'none' });
  });

  test('a row inserted inside a vertical merge grows the merge; its anchor row deleted keeps the text', async () => {
    const rows =
      `<a:tr h="370840">${cell('A', { attrs: ' rowSpan="2"' })}${cell('B')}${cell('C')}</a:tr>` +
      `<a:tr h="370840">${cell('', { attrs: ' vMerge="1"' })}${cell('E')}${cell('F')}</a:tr>` +
      `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`;
    const e = await open(deck(rows));
    expectValid(e);
    caretIn(e, 0, 1);
    e.exec({ type: 'insertRow', where: 'below' });
    expect(attr(model(e).tcs[0][0], 'rowSpan')).toBe('3');
    expectValid(e);
    caretIn(e, 0, 1);
    e.exec({ type: 'deleteRow' });
    expect(textOf(child(model(e).tcs[0][0], 'a:txBody'))).toBe('A');
    expect(attr(model(e).tcs[0][0], 'rowSpan')).toBe('2');
    expectValid(e);
    roundTrip(e);
  });
});

describe('merge, split, and formatting (editor-ui-parity 4.3)', () => {
  test('merge a 2 × 2 range: merge markers, the text of every cell kept in order', async () => {
    const e = await open();
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 1, column: 1 },
    });
    expect(e.exec({ type: 'mergeCells' }).ok).toBe(true);
    const m = model(e);
    expect([attr(m.tcs[0][0], 'gridSpan'), attr(m.tcs[0][0], 'rowSpan')]).toEqual(['2', '2']);
    expect(attr(m.tcs[1][0], 'vMerge')).toBe('1');
    expect([attr(m.tcs[1][1], 'hMerge'), attr(m.tcs[1][1], 'vMerge')]).toEqual(['1', '1']);
    expect(children(child(m.tcs[0][0], 'a:txBody'), 'a:p').map((p) => textOf(p))).toEqual([
      'A',
      'B',
      'D',
      'E',
    ]);
    // The caret sits after the merged text: a key adds to it, it does not prefix it.
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 0, column: 0 },
      focus: { paragraph: 3, offset: 1 },
    });
    expectValid(e);
    roundTrip(e);
    caretIn(e, 0, 0);
    expect(e.can({ type: 'mergeCells' }).ok).toBe(false);
    // Split it back along its grid lines.
    expect(e.exec({ type: 'splitCell', rows: 2, cols: 2 }).ok).toBe(true);
    expect(
      model(e)
        .anchors.flat()
        .filter((a, i, all) => all.findIndex((b) => b.row === a.row && b.column === a.column) === i)
    ).toHaveLength(9);
    expectValid(e);
  });

  test('split one cell into new grid columns and rows; the cells beside it span them', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    expect(e.exec({ type: 'splitCell', rows: 2, cols: 2 }).ok).toBe(true);
    const m = model(e);
    expect([m.nR, m.nC]).toEqual([4, 4]);
    expect(m.gridCols.map((g) => attr(g, 'w'))).toEqual(['914400', '457200', '457200', '914400']);
    expect(attr(m.tcs[0][1], 'gridSpan')).toBe('2');
    expect(attr(m.tcs[1][0], 'rowSpan')).toBe('2');
    expectValid(e);
    roundTrip(e);
  });

  test('cell shading, borders, margins, anchor, style, and options', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    e.exec({ type: 'setCellFill', color: '#FF0000' });
    const tcPr = () => child(model(e).tcs[1][1], 'a:tcPr')!;
    expect(attr(find(tcPr(), 'a:srgbClr'), 'val')).toBe('FF0000');
    e.exec({ type: 'setTableBorders', scope: 'outside', spec: { color: '#0000FF', width: 2 } });
    for (const side of ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'])
      expect(attr(child(tcPr(), side), 'w')).toBe('25400');
    // The neighbor across each edge agrees.
    expect(attr(child(child(model(e).tcs[1][0], 'a:tcPr'), 'a:lnR'), 'w')).toBe('25400');
    expect(elements(tcPr()).map((c) => c.local)).toEqual(['lnL', 'lnR', 'lnT', 'lnB', 'solidFill']);
    e.exec({ type: 'setCellMargins', margins: { left: 0, top: 91_440 } });
    expect([attr(tcPr(), 'marL'), attr(tcPr(), 'marT')]).toEqual(['0', '91440']);
    e.exec({ type: 'setVerticalAlign', anchor: 'ctr' });
    expect(attr(tcPr(), 'anchor')).toBe('ctr');
    expect(e.snapshot().format?.verticalAlign).toBe('ctr');
    e.exec({ type: 'toggleTableStyleOption', option: 'bandRow' });
    e.exec({ type: 'toggleTableStyleOption', option: 'firstCol' });
    expect(e.snapshot().table?.options).toEqual(['firstRow', 'firstCol']);
    const other = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}';
    e.exec({ type: 'setTableStyle', styleId: other });
    expect(textOf(find(root(e), 'a:tableStyleId'))).toBe(other);
    expect(e.can({ type: 'setTableStyle', styleId: 'x' }).ok).toBe(false);
    e.exec({ type: 'setTableBorders', scope: 'none' });
    expect(child(child(tcPr(), 'a:lnL'), 'a:noFill')).toBeTruthy();
    roundTrip(e);
  });

  test('the border pen: a target pick draws the pen, a pen pick redraws the last target', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    const tcPr = () => child(model(e).tcs[1][1], 'a:tcPr')!;
    expect(toolbarCommandState(e, 'table.borderWidth').value).toBe('1');
    // A pen pick before any target pick arms the pen and writes nothing.
    const before = xml(e);
    expect(runToolbarCommand(e, 'table.borderColor', '#00FF00').ok).toBe(true);
    expect(xml(e)).toBe(before);
    expect(e.snapshot().canUndo).toBe(false);
    // A target pick draws the armed pen there; then a pen pick redraws that target.
    expect(runToolbarCommand(e, 'table.borderTarget', 'outside').ok).toBe(true);
    expect(attr(find(child(tcPr(), 'a:lnT')!, 'a:srgbClr'), 'val')).toBe('00FF00');
    expect(runToolbarCommand(e, 'table.borderWidth', '3').ok).toBe(true);
    expect(attr(child(tcPr(), 'a:lnL'), 'w')).toBe('38100');
    expect(toolbarCommandState(e, 'table.borderWidth').value).toBe('3');
    expect(runToolbarCommand(e, 'table.borderStyle', 'dash').ok).toBe(true);
    expect(attr(find(child(tcPr(), 'a:lnB')!, 'a:prstDash'), 'val')).toBe('dash');
    expect(toolbarCommandState(e, 'table.borderStyle').value).toBe('dash');
    // The pen keeps its color: the outside borders stay green.
    expect(attr(find(child(tcPr(), 'a:lnB')!, 'a:srgbClr'), 'val')).toBe('00FF00');
    // An invalid pen value changes nothing.
    expect(runToolbarCommand(e, 'table.borderWidth', '-1' as never).ok).toBe(false);
    expect(toolbarCommandState(e, 'table.borderWidth').value).toBe('3');
    // In another cell, a pen pick arms only: the pen never draws where the user did not pick.
    caretIn(e, 0, 0);
    const moved = xml(e);
    expect(runToolbarCommand(e, 'table.borderColor', '#FF0000').ok).toBe(true);
    expect(xml(e)).toBe(moved);
    expect(runToolbarCommand(e, 'table.borderTarget', 'outside').ok).toBe(true);
    expect(runToolbarCommand(e, 'table.borderColor', '#00FF00').ok).toBe(true);
    // The same range, selected from the other corner, is the same cells: a pick redraws.
    const cells = (r0: number, c0: number, r1: number, c1: number) =>
      e.setSelection({
        kind: 'cells',
        nodeId: frameId(e),
        anchor: { row: r0, column: c0 },
        focus: { row: r1, column: c1 },
      });
    cells(0, 0, 1, 1);
    expect(runToolbarCommand(e, 'table.borderTarget', 'outside').ok).toBe(true);
    cells(1, 1, 0, 0);
    const drawn = xml(e);
    expect(runToolbarCommand(e, 'table.borderColor', '#0000FF').ok).toBe(true);
    expect(xml(e)).not.toBe(drawn);
    caretIn(e, 1, 1);
    // After No Border, a pen pick arms again and draws nothing.
    expect(runToolbarCommand(e, 'table.borderTarget', 'none').ok).toBe(true);
    const cleared = xml(e);
    expect(runToolbarCommand(e, 'table.borderColor', '#0000FF').ok).toBe(true);
    expect(xml(e)).toBe(cleared);
    // In view mode a pen pick is refused, as a draw would be, and the pen stays.
    e.setMode('view');
    expect(runToolbarCommand(e, 'table.borderWidth', '6').ok).toBe(false);
    expect(toolbarCommandState(e, 'table.borderWidth').value).toBe('3');
    e.setMode('edit');
    expectValid(e);
    roundTrip(e);
  });

  test('row height, column width, distribute, and a grid drag as one undo step', async () => {
    const e = await open(deck(undefined, [IN, 2 * IN, IN]));
    caretIn(e, 0, 0);
    e.exec({ type: 'setRowHeight', height: IN });
    e.exec({ type: 'setColumnWidth', width: IN / 2 });
    expect(attr(model(e).rows[0], 'h')).toBe(String(IN));
    expect(attr(model(e).gridCols[0], 'w')).toBe(String(IN / 2));
    e.exec({ type: 'distributeColumns' });
    expect(model(e).gridCols.map((g) => attr(g, 'w'))).toEqual(
      Array(3).fill(String(Math.round((3.5 * IN) / 3)))
    );
    const ext = () => Number(attr(find(find(root(e), 'p:graphicFrame')!, 'a:ext'), 'cx'));
    expect(Math.abs(ext() - 3.5 * IN)).toBeLessThanOrEqual(2);
    const historyGroup = e.beginHistoryGroup();
    e.exec({ type: 'setTableGrid', columns: [IN, IN, IN] }, { historyGroup });
    e.exec({ type: 'setTableGrid', columns: [2 * IN, IN, IN] }, { historyGroup });
    historyGroup.end();
    e.exec({ type: 'undo' });
    expect(model(e).gridCols.map((g) => attr(g, 'w'))).toEqual(
      Array(3).fill(String(Math.round((3.5 * IN) / 3)))
    );
    expect(e.can({ type: 'setTableGrid', columns: [IN] }).ok).toBe(false);
    roundTrip(e);
  });

  test('Shape Fill shades the cells of a table selection', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    expect(e.exec({ type: 'setShapeFill', color: '#00FF00' }).ok).toBe(true);
    expect(
      model(e)
        .tcs.flat()
        .every((tc) => !!find(child(tc, 'a:tcPr'), 'a:srgbClr'))
    ).toBe(true);
  });

  test('table commands refuse without a table', async () => {
    const e = await open();
    e.setSelection({ kind: 'none' });
    expect(e.can({ type: 'insertRow', where: 'above' })).toEqual({
      ok: false,
      reason: 'disabled.noSelection',
    });
    const box = elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'sp')!.id;
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    expect(e.can({ type: 'mergeCells' })).toEqual({ ok: false, reason: 'disabled.noTable' });
    expect(e.snapshot().table).toBeNull();
  });

  test('a command type the engine does not know is refused, not thrown', async () => {
    const e = await open();
    const unknown = { type: 'insertTableRow' } as unknown as Parameters<typeof e.exec>[0];
    expect(e.exec(unknown)).toMatchObject({ ok: false, reason: 'errors.unsupported' });
    expect(e.can(unknown)).toMatchObject({ ok: false, reason: 'errors.unsupported' });
  });
});

describe('repair of a table that is not valid', () => {
  // A short row (two cells for three grid columns) and an orphan hMerge that no span claims.
  const broken = () =>
    deck(
      `<a:tr h="370840">${cell('A')}${cell('B')}${cell('C')}</a:tr>` +
        `<a:tr h="370840">${cell('D')}${cell('E', { attrs: ' hMerge="1"' })}</a:tr>` +
        `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`
    );

  test('save writes a copy with whole rows and leaves the store as it is', async () => {
    const e = await open(broken());
    const before = semanticDigest(e.session!.store.package);
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    expect(semanticDigest(e.session!.store.package)).toBe(before);
    if (!r.ok) return;
    const again = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await again.load(writePptx(r.pkg))).ok).toBe(true);
    expect(model(again).tcs.map((row) => row.length)).toEqual([3, 3, 3]);
    expect(grid(again)[1]).toEqual(['D', 'E', '']);
    // A save keeps the merge markers as the file has them: it never changes what a cell shows.
    expect(attr(model(again).tcs[1][1], 'hMerge')).toBe('1');
  });

  test('a save keeps the hidden text of a covered cell that the file brought', async () => {
    const rows =
      `<a:tr h="370840">${cell('Total', { attrs: ' gridSpan="2"' })}` +
      `${cell('Total', { attrs: ' hMerge="1"' })}${cell('C')}</a:tr>`;
    const e = await open(deck(rows));
    roundTrip(e);
    const r = readPptx(e.save()!);
    expect(r.ok && canonicalPackageFingerprint(r.pkg)).toBe(
      canonicalPackageFingerprint(e.session!.store.package)
    );
    expect(grid(e)[0]).toEqual(['Total', 'Total', 'C']);
  });

  test('a split shows the hidden text of a covered cell that the file brought', async () => {
    const hidden =
      '<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>' +
      '<a:t>secret</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>';
    const rows = `<a:tr h="370840">${cell('A', { attrs: ' gridSpan="2"' })}${hidden}${cell('C')}</a:tr>`;
    const e = await open(deck(rows));
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(textOf(find(xmlPartOf(r.pkg, part(e))!.root, 'a:tbl'))).toContain('secret');
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'splitCell', rows: 1, cols: 2 }).ok).toBe(true);
    expect(grid(e)[0]).toEqual(['A', 'secret', 'C']);
  });

  test('a table command repairs the table it edits, in one undo step', async () => {
    const e = await open(broken());
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    expectValid(e);
    expect(model(e).nR).toBe(4);
    roundTrip(e);
    // Undo takes the row and the repair back together: the short row is short again.
    e.exec({ type: 'undo' });
    const tbl = find(root(e), 'a:tbl')!;
    expect(children(children(tbl, 'a:tr')[1], 'a:tc')).toHaveLength(2);
  });
});

test('a covered cell without a text body gets one when a split shows it', async () => {
  const e = await open(
    deck(
      `<a:tr h="370840">${cell('AB', { attrs: ' gridSpan="2"' })}<a:tc hMerge="1"><a:tcPr/></a:tc>${cell('C')}</a:tr>` +
        `<a:tr h="370840">${cell('D')}${cell('E')}${cell('F')}</a:tr>` +
        `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`
    )
  );
  caretIn(e, 0, 0);
  expect(e.exec({ type: 'splitCell', rows: 1, cols: 2 }).ok).toBe(true);
  expectValid(e);
  caretIn(e, 0, 1);
  expect(e.input('text', 'x').ok).toBe(true);
  expect(grid(e)[0]).toEqual(['AB', 'x', 'C']);
  roundTrip(e);
});

describe('structure edits through merges, and repair edge cases', () => {
  // Row 0: "AB" merged over columns 0 and 1, then "C".
  const merged = () =>
    deck(
      `<a:tr h="370840">${cell('AB', { attrs: ' gridSpan="2"' })}${cell('', { attrs: ' hMerge="1"' })}${cell('C')}</a:tr>` +
        `<a:tr h="370840">${cell('D')}${cell('E')}${cell('F')}</a:tr>` +
        `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`
    );
  const cellsOf = (e: PptxEditorInstance, r: number): Element[] =>
    children(children(find(root(e), 'a:tbl')!, 'a:tr')[r], 'a:tc');

  test('Select column under a merged header selects one column; Delete Columns deletes it', async () => {
    const e = await open(merged());
    caretIn(e, 1, 1);
    expect(e.exec({ type: 'selectTableRegion', region: 'column' }).ok).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      kind: 'cells',
      anchor: { row: 0, column: 1 },
      focus: { row: 2, column: 1 },
    });
    expect(e.exec({ type: 'deleteColumn' }).ok).toBe(true);
    // One column goes; the merged header shrinks to the column that is left.
    expect(grid(e)).toEqual([
      ['AB', 'C'],
      ['D', 'F'],
      ['G', 'I'],
    ]);
    expectValid(e);
    roundTrip(e);
  });

  test('a column inserted inside a horizontal merge widens the merge', async () => {
    const e = await open(merged());
    caretIn(e, 1, 0);
    expect(e.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    expectValid(e);
    expect(attr(cellsOf(e, 0)[0], 'gridSpan')).toBe('3');
    expect(grid(e)[1]).toEqual(['D', '', 'E', 'F']);
    roundTrip(e);
  });

  test('deleting the first column of a merge moves its text to the cell that stays', async () => {
    const e = await open(merged());
    caretIn(e, 1, 0);
    expect(e.exec({ type: 'deleteColumn' }).ok).toBe(true);
    expectValid(e);
    expect(grid(e)[0]).toEqual(['AB', 'C']);
    roundTrip(e);
  });

  test('a moved merge anchor keeps the cell child order when the cell had no text body', async () => {
    const e = await open(
      deck(
        `<a:tr h="370840">${cell('AD', { attrs: ' rowSpan="2"' })}${cell('B')}${cell('C')}</a:tr>` +
          `<a:tr h="370840"><a:tc vMerge="1"><a:tcPr/></a:tc>${cell('E')}${cell('F')}</a:tr>` +
          `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`
      )
    );
    caretIn(e, 0, 1);
    expect(e.exec({ type: 'deleteRow' }).ok).toBe(true);
    expectValid(e);
    expect(elements(cellsOf(e, 0)[0]).map((c) => c.local)).toEqual(['txBody', 'tcPr']);
    expect(grid(e)[0]).toEqual(['AD', 'E', 'F']);
    roundTrip(e);
  });

  test('a long row: blank cells past the grid go; a cell with text gets a grid column', async () => {
    const extra = (last: string) =>
      deck(
        `<a:tr h="370840">${cell('A')}${cell('B')}${cell('C')}${cell(last)}</a:tr>` +
          `<a:tr h="370840">${cell('D')}${cell('E')}${cell('F')}</a:tr>`
      );
    const blank = await open(extra(''));
    caretIn(blank, 0, 0);
    // `can` repairs in its probe too, so it agrees with `exec`.
    expect(blank.can({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    expect(blank.exec({ type: 'distributeColumns' }).ok).toBe(true);
    expect(model(blank).nC).toBe(3);
    expect(model(blank).gridCols.map((g) => attr(g, 'w'))).toEqual(['914400', '914400', '914400']);
    expectValid(blank);
    const text = await open(extra('X'));
    caretIn(text, 0, 0);
    expect(text.exec({ type: 'distributeColumns' }).ok).toBe(true);
    expect(model(text).nC).toBe(4);
    expect(grid(text)[0]).toEqual(['A', 'B', 'C', 'X']);
    expectValid(text);
    roundTrip(text);
  });

  test('a table larger than the editor makes is not repaired, and saves as it is', async () => {
    const wide = Array.from({ length: 80 }, (_, i) => cell(String(i))).join('');
    const e = await open(deck(`<a:tr h="370840">${wide}</a:tr>`));
    const r = readPptx(e.save()!);
    expect(r.ok && semanticDigest(r.pkg)).toBe(semanticDigest(e.session!.store.package));
  });
});

describe('round 3: keys on a cell range, cells without a body, style ids', () => {
  const range = (e: PptxEditorInstance, r0: number, c0: number, r1: number, c1: number) =>
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: r0, column: c0 },
      focus: { row: r1, column: c1 },
    });

  test('Tab, arrows, and Shift+arrows on a cell range keep the focus in the table', async () => {
    const e = await open();
    range(e, 0, 0, 0, 1);
    expect(press(e, 'Tab')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'text', cell: { row: 0, column: 2 } });
    range(e, 0, 0, 0, 1);
    expect(press(e, 'ArrowDown', true)).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'cells', focus: { row: 1, column: 1 } });
    expect(press(e, 'ArrowRight')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'text', cell: { row: 1, column: 2 } });
    // At the table's edge the key is still handled: the focus does not leave the table.
    range(e, 2, 2, 2, 2);
    expect(press(e, 'ArrowDown')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'cells' });
  });

  test('typing on a cell range is refused with a reason for a range', async () => {
    const e = await open();
    range(e, 0, 0, 1, 1);
    expect(e.input('text', 'x')).toMatchObject({ ok: false, reason: 'disabled.typeInOneCell' });
  });

  test('typing in a drawn cell without a text body gives it one', async () => {
    const e = await open(
      deck(
        `<a:tr h="370840">${cell('A')}<a:tc><a:tcPr/></a:tc>${cell('C')}</a:tr>` +
          `<a:tr h="370840">${cell('D')}${cell('E')}${cell('F')}</a:tr>` +
          `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`
      )
    );
    caretIn(e, 0, 1);
    expect(e.input('text', 'B').ok).toBe(true);
    expect(grid(e)[0]).toEqual(['A', 'B', 'C']);
    roundTrip(e);
  });

  test('a table style id must be an upper-case GUID, as ST_Guid allows', async () => {
    const e = await open();
    caretIn(e, 0, 0);
    expect(
      e.exec({ type: 'setTableStyle', styleId: '{073a0daa-6af3-43ab-8588-cec1d06c72b9}' }).ok
    ).toBe(false);
    expect(
      e.exec({ type: 'setTableStyle', styleId: '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}' }).ok
    ).toBe(true);
  });
});

describe('a format armed at a caret in a cell', () => {
  test('Bold at a word edge in a cell arms, and the typed text takes it in that cell', async () => {
    const e = await open();
    caretIn(e, 1, 1, 1);
    const before = grid(e);
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(grid(e)).toEqual(before);
    expect(e.snapshot().format?.bold).toBe(true);
    expect(e.input('text', 'x').ok).toBe(true);
    expect(grid(e)[1]).toEqual(['D', 'Ex', 'F']);
    const tc = children(children(find(root(e), 'a:tbl')!, 'a:tr')[1], 'a:tc')[1];
    const xRun = elements(find(tc, 'a:p')!).filter((c) => c.local === 'r')[1]!;
    expect(attr(child(xRun, 'a:rPr')!, 'b')).toBe('1');
    // One undo removes the text and its format.
    e.exec({ type: 'undo' });
    expect(grid(e)).toEqual(before);
  });
});

describe('the format read of a selected table or cell range', () => {
  test('reads every cell the command writes: bullets, paragraphs, and a mixed font', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    const numbers = { kind: 'number', scheme: 'arabicPeriod' } as const;
    expect(e.exec({ type: 'setBulletFormat', format: numbers }).ok).toBe(true);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({ kind: 'number' });
    expect(e.snapshot().format?.paragraph).toBeDefined();
    // Bold in one cell: a range over two cells reads the font as mixed.
    caretIn(e, 0, 0);
    const sel = e.snapshot().selection as Extract<EditorSelection, { kind: 'text' }>;
    e.setSelection({ ...sel, focus: { paragraph: 0, offset: 1 } });
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(e.snapshot().format?.font?.disagrees.bold).toBe(true);
    expect(e.snapshot().format?.bulletFormat).toMatchObject({ kind: 'number' });
  });
});

describe('round 4: repair order, a second text body, and chrome state', () => {
  const lines = (w: string) =>
    `<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>`.replace(/^/, `<a:${w} w="12700">`) +
    `</a:${w}>`;

  test('the repair puts a:tcPr children in schema order and keeps text of a second body', async () => {
    const e = await open(
      deck(
        `<a:tr h="370840">${cell('A', { tcPr: lines('lnT') + lines('lnL') })}` +
          `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>X</a:t></a:r></a:p></a:txBody>` +
          `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Y</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
          `${cell('C')}</a:tr>` +
          `<a:tr h="370840">${cell('D')}${cell('E')}${cell('F')}</a:tr>`
      )
    );
    caretIn(e, 1, 0);
    // A range: Bold at a bare caret on a word edge only arms for the next typed text.
    const sel = e.snapshot().selection as Extract<EditorSelection, { kind: 'text' }>;
    e.setSelection({ ...sel, focus: { paragraph: 0, offset: 1 } });
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    const cells = children(children(find(root(e), 'a:tbl')!, 'a:tr')[0], 'a:tc');
    expect(elements(child(cells[0], 'a:tcPr')!).map((c) => c.local)).toEqual(['lnL', 'lnT']);
    expect(elements(cells[1]).map((c) => c.local)).toEqual(['txBody', 'tcPr']);
    expect(grid(e)[0][1]).toBe('XY');
    roundTrip(e);
  });

  test('an inline a:tableStyle stays when a style id sits beside it', async () => {
    const inline =
      '<a:tableStyle styleId="{11111111-2222-3333-4444-555555555555}" styleName="Inline"><a:wholeTbl><a:tcTxStyle/></a:wholeTbl></a:tableStyle>';
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: table({
              id: 2,
              x: IN,
              y: IN,
              cx: 2 * IN,
              cy: 370_840,
              columns: [IN, IN],
              tblPr: `<a:tblPr><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId>${inline}</a:tblPr>`,
              rows: `<a:tr h="370840">${cell('A')}${cell('B')}</a:tr>`,
            }),
          },
        ],
      })
    );
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const slide = xmlPartOf(r.pkg, part(e))!.root;
    expect(elements(find(slide, 'a:tblPr')!).map((c) => c.local)).toEqual(['tableStyle']);
  });

  test('Split Cells is enabled when any preset can run', async () => {
    const e = await open(
      deck(
        `<a:tr h="370840">${cell('A', { attrs: ' rowSpan="3"' })}${cell('B')}${cell('C')}</a:tr>` +
          `<a:tr h="370840">${cell('', { attrs: ' vMerge="1"' })}${cell('E')}${cell('F')}</a:tr>` +
          `<a:tr h="370840">${cell('', { attrs: ' vMerge="1"' })}${cell('H')}${cell('I')}</a:tr>`
      )
    );
    caretIn(e, 0, 0);
    expect(e.can({ type: 'splitCell', rows: 2, cols: 1 }).ok).toBe(false);
    expect(toolbarCommandState(e, 'table.splitCell').enabled).toBe(true);
  });

  test('Shift+arrow back to the range anchor selects that cell text', async () => {
    const e = await open();
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(press(e, 'ArrowLeft', true)).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'text', cell: { row: 0, column: 0 } });
  });

  test('Shift+arrow at the edge of the cell text starts a cell range', async () => {
    const e = await open();
    // Inside the text, Shift+arrow extends the text selection as usual.
    caretIn(e, 1, 1, 0);
    expect(press(e, 'ArrowRight', true)).toBe(false);
    // At the end of the text, it selects this cell and the next one.
    caretIn(e, 1, 1, 1);
    expect(press(e, 'ArrowRight', true)).toBe(true);
    expect(e.snapshot().selection).toEqual({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 1, column: 1 },
      focus: { row: 1, column: 2 },
    });
    // The range then grows down with Shift+ArrowDown.
    expect(press(e, 'ArrowDown', true)).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'cells', focus: { row: 2, column: 2 } });
    // At the table's edge, nothing to extend to: the key is left to text editing.
    caretIn(e, 0, 2, 1);
    expect(press(e, 'ArrowRight', true)).toBe(false);
  });
});

describe('size and split limits', () => {
  test('a table cannot grow past 75 columns, with its own reason', async () => {
    const columns = Array.from({ length: 75 }, () => IN / 10);
    const row = `<a:tr h="370840">${columns.map(() => cell('x')).join('')}</a:tr>`;
    const e = await open(deck(row, columns));
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'insertColumn', where: 'right' })).toMatchObject({
      ok: false,
      reason: 'disabled.tableTooLarge',
    });
    expect(e.exec({ type: 'splitCell', rows: 1, cols: 2 })).toMatchObject({
      ok: false,
      reason: 'disabled.tableTooLarge',
    });
  });

  test('a merged cell splits only into parts that divide it, with its own reason', async () => {
    const e = await open();
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(e.exec({ type: 'mergeCells' }).ok).toBe(true);
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'splitCell', rows: 1, cols: 3 })).toMatchObject({
      ok: false,
      reason: 'disabled.splitUneven',
    });
  });
});

test('a merged cell over 2 rows does not split into 3 rows, and the refusal changes nothing', async () => {
  const e = await open(
    deck(
      `<a:tr h="370840">${cell('AD', { attrs: ' rowSpan="2"' })}${cell('B')}${cell('C')}</a:tr>` +
        `<a:tr h="370840">${cell('', { attrs: ' vMerge="1"' })}${cell('E')}${cell('F')}</a:tr>` +
        `<a:tr h="370840">${cell('G')}${cell('H')}${cell('I')}</a:tr>`
    )
  );
  const before = semanticDigest(e.session!.store.package);
  caretIn(e, 0, 0);
  expect(e.exec({ type: 'splitCell', rows: 3, cols: 1 })).toMatchObject({
    ok: false,
    reason: 'disabled.splitUneven',
  });
  expect(semanticDigest(e.session!.store.package)).toBe(before);
});
