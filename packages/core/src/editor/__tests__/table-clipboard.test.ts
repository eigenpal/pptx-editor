// A cell range on the clipboard (table-clipboard.ts) and the `pasteCells` command: copy writes
// tab-separated text and the cells format; paste fills cells from the selection's top-left cell.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { child, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import { DEFAULT_TABLE_STYLE_ID } from '../../cascade/index.ts';
import { createPptxEditor, toolbarCommandState, type PptxEditorInstance } from '../index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { readTable } from '../../automation/commands/table-model.ts';
import {
  CELLS_MIME,
  cellRangeTexts,
  parseCellsJson,
  parseTabSeparated,
  pastedCells,
  toTabSeparated,
} from '../table-clipboard.ts';

const IN = 914_400;

/** A 3 × 3 table: "A B C", "D E F", "G H I". `merge` spans row 1's first two cells. */
function deck(merge = false): Uint8Array {
  const rows = ['ABC', 'DEF', 'GHI'].map((letters, r) => {
    const cells =
      merge && r === 1
        ? '<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>DE</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>' +
          cell('F')
        : letters
            .split('')
            .map((t) => cell(t))
            .join('');
    return `<a:tr h="370840">${cells}</a:tr>`;
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
            cx: 3 * IN,
            cy: 3 * 370_840,
            columns: [IN, IN, IN],
            tblPr: `<a:tblPr firstRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
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
const frameId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
const grid = (e: PptxEditorInstance): string[][] =>
  readTable(new OpBatch(e.session!.store.package), part(e), frameId(e)).tcs.map((row) =>
    row.map((tc) => textOf(child(tc, 'a:txBody')))
  );
const range = (e: PptxEditorInstance, r0: number, c0: number, r1: number, c1: number) =>
  e.setSelection({
    kind: 'cells',
    nodeId: frameId(e),
    anchor: { row: r0, column: c0 },
    focus: { row: r1, column: c1 },
  });
const caretIn = (e: PptxEditorInstance, row: number, column: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });

describe('tab-separated text', () => {
  test('quotes cells with tabs, line breaks, and quotes, and reads them back', () => {
    const rows = [
      ['a', 'b\tc', 'say "hi"'],
      ['two\nlines', '', 'z'],
    ];
    const text = toTabSeparated(rows);
    expect(text).toBe('a\t"b\tc"\t"say ""hi"""\n"two\nlines"\t\tz');
    expect(parseTabSeparated(text)).toEqual(rows);
  });

  test('reads CRLF lines, and one trailing line break is not a row', () => {
    expect(parseTabSeparated('1\t2\r\n3\t4\r\n')).toEqual([
      ['1', '2'],
      ['3', '4'],
    ]);
    expect(parseTabSeparated('only')).toEqual([['only']]);
  });

  test('the cells format must be rows of strings', () => {
    expect(parseCellsJson('[["a","b"],["c"]]')).toEqual([['a', 'b'], ['c']]);
    expect(parseCellsJson('[["a",1]]')).toBeNull();
    expect(parseCellsJson('{"rows":[]}')).toBeNull();
    expect(parseCellsJson('not json')).toBeNull();
  });

  test('reads at most 75 rows of 75 cells, and no more than 4 MB of text', () => {
    const rows = parseTabSeparated(Array.from({ length: 200 }, () => 'a\tb').join('\n'));
    expect(rows).toHaveLength(75);
    expect(parseTabSeparated(Array(100).fill('x').join('\t'))[0]).toHaveLength(75);
    expect(parseCellsJson(`[["${'x'.repeat(4_000_001)}"]]`)).toBeNull();
  });
});

describe('copy a cell range', () => {
  test('the range reads row by row; a covered position is empty', async () => {
    const e = await open(deck(true));
    range(e, 0, 0, 1, 2);
    expect(cellRangeTexts(e.session!.layout(0), e.snapshot().selection)).toEqual([
      ['A', 'B', 'C'],
      ['DE', '', 'F'],
    ]);
  });

  test('Copy and Cut are enabled on a cell range; Paste reads the system clipboard', async () => {
    const e = await open();
    range(e, 0, 0, 1, 1);
    for (const id of ['edit.copy', 'edit.cut', 'edit.paste'] as const)
      expect(toolbarCommandState(e, id).enabled).toBe(true);
  });
});

describe('paste into cells', () => {
  test('pasteCells fills cells from the caret, leaves out text past the edge, and undoes once', async () => {
    const e = await open();
    caretIn(e, 1, 1);
    const r = e.exec({
      type: 'pasteCells',
      texts: [
        ['x', 'y', 'past the edge'],
        ['z', 'w'],
      ],
    });
    expect(r.ok).toBe(true);
    expect(grid(e)).toEqual([
      ['A', 'B', 'C'],
      ['D', 'x', 'y'],
      ['G', 'z', 'w'],
    ]);
    // The filled cells are the selection.
    expect(e.snapshot().selection).toMatchObject({
      kind: 'cells',
      anchor: { row: 1, column: 1 },
      focus: { row: 2, column: 2 },
    });
    e.exec({ type: 'undo' });
    expect(grid(e)[1]).toEqual(['D', 'E', 'F']);
  });

  test('a cell keeps the format of its first run; lines become paragraphs', async () => {
    const e = await open(deck(true));
    caretIn(e, 1, 0);
    expect(e.exec({ type: 'pasteCells', texts: [['one\ntwo', 'skipped', 'G2']] }).ok).toBe(true);
    const m = readTable(new OpBatch(e.session!.store.package), part(e), frameId(e));
    const body = child(m.tcs[1][0], 'a:txBody')!;
    const paragraphs = elements(body).filter((n) => n.local === 'p');
    expect(paragraphs).toHaveLength(2);
    // The covered position is skipped; the third text goes into column 2.
    expect(grid(e)[1][2]).toBe('G2');
    const rPr = find(paragraphs[1], 'a:rPr');
    expect(rPr && rPr.attrs.find((a) => a.local === 'b')?.value).toBe('1');
    // The saved file reopens with the same text.
    const saved = await e.save();
    const again = await open(saved!);
    expect(grid(again)[1][2]).toBe('G2');
  });

  test('a paste taller or wider than the table fills what fits', async () => {
    const e = await open();
    caretIn(e, 1, 0);
    const texts = Array.from({ length: 20 }, (_, i) => Array.from({ length: 9 }, () => `r${i}`));
    expect(e.exec({ type: 'pasteCells', texts }).ok).toBe(true);
    expect(grid(e)).toEqual([
      ['A', 'B', 'C'],
      ['r0', 'r0', 'r0'],
      ['r1', 'r1', 'r1'],
    ]);
  });

  test('pasteCells refuses texts that are not rows of strings', async () => {
    const e = await open();
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'pasteCells', texts: [] }).ok).toBe(false);
    expect(e.exec({ type: 'pasteCells', texts: [[1 as unknown as string]] }).ok).toBe(false);
  });

  test('what a paste fills: a copied range, tab-separated text, or any text on a range', async () => {
    const e = await open();
    caretIn(e, 0, 0);
    const sel = e.snapshot().selection;
    expect(pastedCells(sel, { cells: '[["a"]]', text: 'a' })).toEqual([['a']]);
    expect(pastedCells(sel, { text: 'a\tb' })).toEqual([['a', 'b']]);
    // Plain text into a caret is the cell's own text paste.
    expect(pastedCells(sel, { text: 'plain\ntext' })).toBeNull();
    range(e, 0, 0, 1, 1);
    expect(pastedCells(e.snapshot().selection, { text: 'plain' })).toEqual([['plain']]);
    // Outside a table, never.
    e.setSelection({ kind: 'none' });
    expect(pastedCells(e.snapshot().selection, { text: 'a\tb' })).toBeNull();
    expect(CELLS_MIME).toContain('cells');
  });
});
