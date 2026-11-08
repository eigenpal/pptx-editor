// Table editing: review fixes (links in new cells, frame resize, inherited
// bullets, group scope, context menu, replace, view mode, the default style).

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
import { DEFAULT_TABLE_STYLE_ID, defaultTableStyleId } from '../../cascade/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { applyTableKey, tableKeyAction } from '../table-keys.ts';
import { selectionScope } from '../../automation/commands/placement.ts';
import { contextMenuSelection } from '../surface-context-menu.ts';
import { readTable, type TableModel } from '../../automation/commands/table-model.ts';

const IN = 914_400;

/** A 3 × 3 table: row 0 "A B C", row 1 "D E F", row 2 "G H I". */
function deck(rows?: string, columns = [IN, IN, IN], rtl = false): Uint8Array {
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
            tblPr: `<a:tblPr${rtl ? ' rtl="1"' : ''} firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
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
const root = (e: PptxEditorInstance): Element => xmlPartOf(e.session!.store.package, part(e))!.root;
const frameId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
const model = (e: PptxEditorInstance): TableModel =>
  readTable(new OpBatch(e.session!.store.package), part(e), frameId(e));
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

describe('review fixes: links, resize, bullets, scope', () => {
  const linked =
    '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1">' +
    '<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>' +
    '</a:rPr><a:t>Link</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>';

  test('a new row takes the format of the cell above, not its link', async () => {
    const e = await open(deck(`<a:tr h="370840">${linked}${cell('B')}${cell('C')}</a:tr>`));
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    const end = find(model(e).tcs[1][0], 'a:endParaRPr')!;
    expect(attr(end, 'b')).toBe('1');
    expect(find(end, 'a:hlinkClick')).toBeUndefined();
    caretIn(e, 1, 0);
    expect(e.input('text', 'x').ok).toBe(true);
    expect(find(model(e).tcs[1][0], 'a:hlinkClick')).toBeUndefined();
  });

  test('a resize of the table frame scales its columns and rows, and stays', async () => {
    const e = await open();
    const id = frameId(e);
    const x = find(find(root(e), 'p:graphicFrame')!, 'a:off')!;
    const ox = Number(attr(x, 'x'));
    const oy = Number(attr(x, 'y'));
    expect(
      e.exec({
        type: 'setTransforms',
        items: [{ nodeId: id, x: ox, y: oy, cx: 6 * IN, cy: 6 * 370_840, rot: 0 }],
      }).ok
    ).toBe(true);
    const cols = children(find(root(e), 'a:tblGrid'), 'a:gridCol').map((g) => attr(g, 'w'));
    expect(cols).toEqual([String(2 * IN), String(2 * IN), String(2 * IN)]);
    expect(model(e).rows.map((r) => attr(r, 'h'))).toEqual(['741680', '741680', '741680']);
    // The next table command keeps the size: the frame is the laid-out table.
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'toggleTableStyleOption', option: 'bandRow' }).ok).toBe(true);
    const ext = find(find(root(e), 'p:graphicFrame')!, 'a:ext')!;
    expect(Number(attr(ext, 'cx'))).toBe(6 * IN);
  });

  test('a drag that sends a box on each move does not compound the row heights', async () => {
    const e = await open();
    const id = frameId(e);
    const off = find(find(root(e), 'p:graphicFrame')!, 'a:off')!;
    const [x, y] = [Number(attr(off, 'x')), Number(attr(off, 'y'))];
    const heights = () => model(e).rows.map((r) => Number(attr(r, 'h')));
    for (const cy of [900_000, 700_000, 500_000, 300_000, 300_000, 300_000, 1_112_520]) {
      const items = [{ nodeId: id, x, y, cx: 3 * IN, cy, rot: 0 }];
      expect(e.exec({ type: 'setTransforms', items }).ok).toBe(true);
    }
    // The last box is the start size: the rows are back to their first minimums.
    expect(heights()).toEqual([370_840, 370_840, 370_840]);
  });

  test('with a row that text holds, a width drag keeps the rows and a smaller box never grows', async () => {
    const e = await open();
    const id = frameId(e);
    caretIn(e, 0, 0, 1);
    for (let i = 0; i < 6; i += 1) e.input('paragraph');
    const box = () => {
      const g = find(root(e), 'p:graphicFrame')!;
      const off = find(g, 'a:off')!;
      const ext = find(g, 'a:ext')!;
      return ['x', 'y']
        .map((k) => Number(attr(off, k)))
        .concat(['cx', 'cy'].map((k) => Number(attr(ext, k))));
    };
    const heights = () => model(e).rows.map((r) => Number(attr(r, 'h')));
    const [x, y, cx, cy] = box();
    const before = heights();
    for (const w of [4 * IN, 3.5 * IN, 4 * IN]) {
      const items = [{ nodeId: id, x, y, cx: w, cy: box()[3], rot: 0 }];
      expect(e.exec({ type: 'setTransforms', items }).ok).toBe(true);
    }
    expect(heights()).toEqual(before);
    for (const h of [cy - 100_000, cy - 200_000, cy - 200_000]) {
      const items = [{ nodeId: id, x, y, cx: box()[2], cy: h, rot: 0 }];
      expect(e.exec({ type: 'setTransforms', items }).ok).toBe(true);
      expect(box()[3]).toBeLessThanOrEqual(cy);
    }
    void cx;
  });

  test('bullets off in a cell whose bullet is inherited writes buNone', async () => {
    const bulleted =
      '<a:tc><a:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="228600" indent="-228600">' +
      '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:lvl1pPr></a:lstStyle>' +
      '<a:p><a:r><a:rPr lang="en-US"/><a:t>Item</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>';
    const e = await open(deck(`<a:tr h="370840">${bulleted}${cell('B')}${cell('C')}</a:tr>`));
    caretIn(e, 0, 0, 1);
    expect(e.snapshot().format?.bullets).toBe(true);
    expect(e.exec({ type: 'toggleBullets' }).ok).toBe(true);
    expect(find(model(e).tcs[0][0], 'a:buNone')).toBeDefined();
  });

  test('a cell range in an entered group keeps the group, and Escape selects the table in it', async () => {
    const tbl = table({
      id: 3,
      x: IN,
      y: IN,
      cx: 3 * IN,
      cy: 370_840,
      columns: [IN, IN, IN],
      tblPr: '<a:tblPr/>',
      rows: `<a:tr h="370840">${cell('A')}${cell('B')}${cell('C')}</a:tr>`,
    });
    const grouped =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      `<p:grpSpPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${3 * IN}" cy="${2 * IN}"/>` +
      `<a:chOff x="${IN}" y="${IN}"/><a:chExt cx="${3 * IN}" cy="${2 * IN}"/></a:xfrm></p:grpSpPr>` +
      tbl +
      shape({ id: 4, x: IN, y: 2 * IN, cx: IN, cy: IN / 2, text: para(run('Box')) }) +
      '</p:grpSp>';
    const e = await open(buildPptx({ slides: [{ shapes: grouped }] }));
    const frame = find(root(e), 'p:graphicFrame')!;
    const group = find(root(e), 'p:grpSp')!;
    e.setSelection({
      kind: 'cells',
      nodeId: frame.id,
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(selectionScope(e.session!.layout(0)!.items, e.snapshot().selection)).toBe(group.id);
    expect(press(e, 'Escape')).toBe(true);
    expect(e.snapshot().selection).toEqual({
      kind: 'shapes',
      nodeIds: [frame.id],
      group: group.id,
    });
  });

  test('a right-click keeps a cell range in the table under the pointer', () => {
    const cells = {
      kind: 'cells',
      nodeId: 'f',
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    } as const;
    expect(contextMenuSelection(cells, 'f', undefined)).toBeNull();
    expect(contextMenuSelection(cells, 'g', undefined)).toEqual({ kind: 'shapes', nodeIds: ['g'] });
  });

  test('replace in a cell re-fits the table frame', async () => {
    const e = await open();
    const before = Number(attr(find(find(root(e), 'p:graphicFrame')!, 'a:ext')!, 'cy'));
    const long = Array.from({ length: 60 }, () => 'word').join(' ');
    expect(e.exec({ type: 'replaceAll', query: 'E', replacement: long }).ok).toBe(true);
    const after = Number(attr(find(find(root(e), 'p:graphicFrame')!, 'a:ext')!, 'cy'));
    expect(after).toBeGreaterThan(before);
  });

  test('view mode turns a cell range into the selected table; sizes under 1 EMU are refused', async () => {
    const e = await open();
    e.setSelection({
      kind: 'cells',
      nodeId: frameId(e),
      anchor: { row: 0, column: 0 },
      focus: { row: 0, column: 1 },
    });
    expect(e.exec({ type: 'setColumnWidth', width: 0.3 }).ok).toBe(false);
    e.setMode('view');
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [frameId(e)] });
  });

  test('a default style the deck does not define gives new tables the built-in default style', async () => {
    const e = await open();
    const batch = new OpBatch(e.session!.store.package);
    const lst = xmlPartOf(batch.pkg, '/ppt/tableStyles.xml')!.root;
    const other = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}';
    batch.push({
      op: 'setAttribute',
      part: '/ppt/tableStyles.xml',
      node: lst.id,
      name: { ns: '', local: 'def' },
      value: other,
    });
    expect(defaultTableStyleId(batch.pkg)).toBe(DEFAULT_TABLE_STYLE_ID);
  });

  test('a width drag on a table whose frame size is stale keeps every row', async () => {
    const tall = cell(Array.from({ length: 8 }, (_, i) => `line ${i}`).join('\n'));
    const e = await open(deck(`<a:tr h="370840">${tall}${cell('B')}${cell('C')}</a:tr>`));
    const g = find(root(e), 'p:graphicFrame')!;
    const off = find(g, 'a:off')!;
    const ext = find(g, 'a:ext')!;
    const rowsBefore = model(e).rows.map((r) => attr(r, 'h'));
    const items = [
      {
        nodeId: frameId(e),
        x: Number(attr(off, 'x')),
        y: Number(attr(off, 'y')),
        cx: 4 * IN,
        cy: Number(attr(ext, 'cy')),
        rot: 0,
      },
    ];
    expect(e.exec({ type: 'setTransforms', items }).ok).toBe(true);
    expect(model(e).rows.map((r) => attr(r, 'h'))).toEqual(rowsBefore);
  });
});

describe('arrow keys in a table that runs right to left', () => {
  test('Left at the edge of a cell moves to the cell on its left, the next one', async () => {
    const e = await open(deck(undefined, undefined, true));
    caretIn(e, 0, 1, 0);
    expect(press(e, 'ArrowLeft')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 0, column: 2 },
      focus: { offset: 0 },
    });
    caretIn(e, 0, 1, 1);
    expect(press(e, 'ArrowRight')).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      cell: { row: 0, column: 0 },
      focus: { offset: 1 },
    });
    // Shift with Left starts a cell range with the cell on the left.
    caretIn(e, 0, 1, 0);
    expect(press(e, 'ArrowLeft', true)).toBe(true);
    expect(e.snapshot().selection).toMatchObject({
      kind: 'cells',
      anchor: { row: 0, column: 1 },
      focus: { row: 0, column: 2 },
    });
  });
});
