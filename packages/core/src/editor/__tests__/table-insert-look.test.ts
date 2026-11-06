// Inserted rows and columns: beside merged cells a new cell takes the look of its block's
// anchor, as a covered cell holds no fill of its own (table-model.ts); after an insert from a
// cell range, the new lines are selected.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, table } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, textOf, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { readTable } from '../../automation/commands/table-model.ts';

const IN = 914_400;
const ORANGE = '<a:solidFill><a:srgbClr val="FF8800"/></a:solidFill>';

/** Row 0: "Half" over two columns, orange, then "C"; rows 1 and 2: plain cells. */
async function open(): Promise<PptxEditorInstance> {
  const rows =
    `<a:tr h="370840">${cell('Half', { attrs: ' gridSpan="2"', tcPr: ORANGE })}` +
    `${cell('', { attrs: ' hMerge="1"' })}${cell('C', { tcPr: ORANGE })}</a:tr>` +
    `<a:tr h="370840">${cell('D', { attrs: ' rowSpan="2"', tcPr: ORANGE })}${cell('E')}${cell('F')}</a:tr>` +
    `<a:tr h="370840">${cell('', { attrs: ' vMerge="1"' })}${cell('H')}${cell('I')}</a:tr>`;
  const bytes = buildPptx({
    slides: [
      {
        shapes: table({
          id: 3,
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: IN,
          columns: [IN, IN, IN],
          tblPr: '<a:tblPr/>',
          rows,
        }),
      },
    ],
  });
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const part = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const frame = (e: PptxEditorInstance) => {
  const root = xmlPartOf(e.session!.store.package, part(e))!.root;
  return elements(find(root, 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
};
const tcs = (e: PptxEditorInstance) =>
  readTable(new OpBatch(e.session!.store.package), part(e), frame(e)).tcs;
const fill = (tc: Parameters<typeof child>[0]) =>
  attr(find(child(tc, 'a:tcPr'), 'a:srgbClr'), 'val');
const caretIn = (e: PptxEditorInstance, row: number, column: number) => {
  const at = { paragraph: 0, offset: 0 };
  e.setSelection({ kind: 'text', nodeId: frame(e), cell: { row, column }, anchor: at, focus: at });
};

describe('inserted cells beside merged cells', () => {
  test('a column beside a merged header takes the header fill, empty', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    expect(e.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    const added = tcs(e)[0][2];
    expect(fill(added)).toBe('FF8800');
    expect(textOf(added)).toBe('');
  });

  test('a row below a tall merged cell takes its fill, empty', async () => {
    const e = await open();
    caretIn(e, 2, 1);
    expect(e.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    const added = tcs(e)[3][0];
    expect(fill(added)).toBe('FF8800');
    expect(textOf(added)).toBe('');
  });

  test('an insert from a cell range selects the new lines; a caret stays in its cell', async () => {
    const e = await open();
    const nodeId = frame(e);
    const column = {
      kind: 'cells',
      nodeId,
      anchor: { row: 0, column: 2 },
      focus: { row: 2, column: 2 },
    } as const;
    e.setSelection(column);
    expect(e.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    expect(e.snapshot().selection).toEqual({
      ...column,
      anchor: { row: 0, column: 3 },
      focus: { row: 2, column: 3 },
    });
    e.setSelection({
      kind: 'cells',
      nodeId,
      anchor: { row: 1, column: 1 },
      focus: { row: 1, column: 2 },
    });
    expect(e.exec({ type: 'insertRow', where: 'above' }).ok).toBe(true);
    const sel = e.snapshot().selection;
    expect(sel.kind === 'cells' && [sel.anchor, sel.focus]).toEqual([
      { row: 1, column: 1 },
      { row: 1, column: 2 },
    ]);
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    expect(e.snapshot().selection).toMatchObject({ kind: 'text', cell: { row: 0, column: 0 } });
  });
});
