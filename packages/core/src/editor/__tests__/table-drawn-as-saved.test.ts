// The canvas draws a table's cells as the save writes them (`tableAsDrawn` in
// `store/package/table-marks.ts`). A cell with two text bodies (a file can hold one; a cell
// holds one `a:txBody`, ECMA-376 Part 1, 21.1.3.16) draws one body with the paragraphs of both,
// as the save joins them; find reads both, and typing or a replace joins them in the store in
// the same undo step. A
// row that concurrent structure edits left short draws the blank cells that the save adds when
// the table has a structure mark, and typing into one adds them. A table without a mark keeps
// its file's cell counts on the canvas.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  children,
  elements,
  find,
  readPptx,
  tableAsDrawn,
  TABLE_MAX_COLUMNS,
  textOf,
  xmlPartOf,
  type Element,
  type PackagePart,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { findMatches } from '../../automation/commands/find.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { markStructureEdit } from '../../automation/commands/table-edit-marks.ts';

const IN = 914_400;
const tr = (...cells: string[]): string => `<a:tr h="370840">${cells.join('')}</a:tr>`;
const body = (text: string): string =>
  `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
/** A cell with two text bodies: `X` in the first, `Y` in the second. */
const TWO_BODIES = `<a:tc>${body('X')}${body('Y')}<a:tcPr/></a:tc>`;

function deck(rows: string, columns = 2): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 3,
          x: IN,
          y: IN,
          cx: columns * IN,
          cy: 2 * 370_840,
          columns: Array.from({ length: columns }, () => IN),
          tblPr: '<a:tblPr/>',
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

const partOf = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const rootOf = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, partOf(e))!.root;
const frameOf = (e: PptxEditorInstance): Element =>
  elements(find(rootOf(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!;

/** Each drawn cell's paragraphs as the canvas lays them out, row by row: `X/Y` for two. */
function drawn(e: PptxEditorInstance): string[] {
  const item = e.session!.layout(0)!.items.find((i) => i.element.kind === 'graphicFrame')!;
  const t = item.element.kind === 'graphicFrame' ? item.element.table! : undefined;
  const out: string[] = [];
  for (const c of t!.cells) {
    if (c.col === 0) out.push('');
    const text = c.text.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('/');
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : text}`;
  }
  return out;
}

/** The text bodies of the cell at row 0, column 1. */
const bodiesOf = (e: PptxEditorInstance): Element[] =>
  children(children(children(find(rootOf(e), 'a:tbl'), 'a:tr')[0], 'a:tc')[1], 'a:txBody');

/** Save `e` and open the save in a new editor. */
async function reopened(e: PptxEditorInstance): Promise<PptxEditorInstance> {
  return open(e.save()!);
}

/** Put a caret at the end of paragraph `paragraph` of cell (`row`, `column`). */
function caret(e: PptxEditorInstance, row: number, column: number, paragraph = 0, offset = 0) {
  e.setSelection({
    kind: 'text',
    nodeId: frameOf(e).id,
    cell: { row, column },
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });
}

/** The slide part of `bytes` with a structure mark on its table, as a peer's command writes it. */
function markedPart(bytes: Uint8Array): PackagePart {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('read');
  const name = [...r.pkg.parts.keys()].find((n) => /slides\/slide1\.xml$/.test(n))!;
  const batch = new OpBatch(r.pkg);
  const tree = find(xmlPartOf(r.pkg, name)!.root, 'p:spTree')!;
  markStructureEdit(batch, name, elements(tree).find((c) => c.local === 'graphicFrame')!.id);
  return batch.pkg.parts.get(name)!;
}

describe('a cell with two text bodies draws the text that the save writes', () => {
  const bytes = deck(tr(cell('A'), TWO_BODIES));

  test('the canvas draws both bodies, find reads both, and the save keeps both in one body', async () => {
    const e = await open(bytes);
    expect(drawn(e)).toEqual(['A|X/Y']);
    expect(findMatches(e.session!.store.package, 'Y')).toHaveLength(1);
    const r = await reopened(e);
    expect(drawn(r)).toEqual(drawn(e));
    expect(bodiesOf(r).map(textOf)).toEqual(['XY']);
  });

  test('replacing text that find reads in the second body joins the bodies in one undo step', async () => {
    const e = await open(bytes);
    expect(e.exec({ type: 'replaceAll', query: 'Y', replacement: 'W' }).ok).toBe(true);
    expect(drawn(e)).toEqual(['A|X/W']);
    expect(bodiesOf(e).map(textOf)).toEqual(['XW']);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(bodiesOf(e).map(textOf)).toEqual(['X', 'Y']);
  });

  test('typing into the second body edits the paragraph the canvas shows, in one undo step', async () => {
    const e = await open(bytes);
    caret(e, 0, 1, 1, 1);
    expect(e.input('text', 'z').ok).toBe(true);
    expect(drawn(e)).toEqual(['A|X/Yz']);
    expect(bodiesOf(e).map(textOf)).toEqual(['XYz']);
    expect(drawn(await reopened(e))).toEqual(['A|X/Yz']);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(drawn(e)).toEqual(['A|X/Y']);
    expect(bodiesOf(e).map(textOf)).toEqual(['X', 'Y']);
  });
});

// Two peers whose table commands each give a cell with no text body an empty one, or one peer's
// command while the other types into the cell, leave a blank body beside another body. A blank
// body adds no empty line (`joinedTextBodies`): the cell draws and saves as one body would.
describe('a blank second text body adds no empty line', () => {
  const BLANK =
    '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>';
  const paragraphCount = (e: PptxEditorInstance) =>
    bodiesOf(e).reduce((n, b) => n + children(b, 'a:p').length, 0);

  test('a blank body after a body with text', async () => {
    const e = await open(deck(tr(cell('A'), `<a:tc>${body('X')}${BLANK}<a:tcPr/></a:tc>`)));
    expect(drawn(e)).toEqual(['A|X']);
    const r = await reopened(e);
    expect(drawn(r)).toEqual(['A|X']);
    expect(paragraphCount(r)).toBe(1);
  });

  test('a blank body before a body with text gives way to it, in the save and in an edit', async () => {
    const e = await open(deck(tr(cell('A'), `<a:tc>${BLANK}${body('Y')}<a:tcPr/></a:tc>`)));
    expect(drawn(e)).toEqual(['A|Y']);
    const r = await reopened(e);
    expect(drawn(r)).toEqual(['A|Y']);
    expect(paragraphCount(r)).toBe(1);
    // The caret's paragraph is the one the canvas shows.
    caret(e, 0, 1, 0, 1);
    expect(e.input('text', 'z').ok).toBe(true);
    expect(drawn(e)).toEqual(['A|Yz']);
    expect(bodiesOf(e).map(textOf)).toEqual(['Yz']);
    expect(paragraphCount(e)).toBe(1);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(bodiesOf(e).map(textOf)).toEqual(['', 'Y']);
  });

  test('a replace of text in the body after a blank first body edits the text find read', async () => {
    const e = await open(deck(tr(cell('A'), `<a:tc>${BLANK}${body('Y')}<a:tcPr/></a:tc>`)));
    expect(e.exec({ type: 'replaceAll', query: 'Y', replacement: 'W' }).ok).toBe(true);
    expect(drawn(e)).toEqual(['A|W']);
    expect(bodiesOf(e).map(textOf)).toEqual(['W']);
  });

  test('two blank bodies draw and save one empty paragraph', async () => {
    const e = await open(deck(tr(cell('A'), `<a:tc>${BLANK}${BLANK}<a:tcPr/></a:tc>`)));
    expect(drawn(e)).toEqual(['A|']);
    const r = await reopened(e);
    expect(drawn(r)).toEqual(['A|']);
    expect(paragraphCount(r)).toBe(1);
  });
});

describe('a short row draws the cells that the save writes', () => {
  const bytes = deck(tr(cell('A'), cell('B')) + tr(cell('C')));

  test('a table without a mark keeps its file cell counts on the canvas', async () => {
    const e = await open(bytes);
    const t = tableAsDrawn(frameOf(e), TABLE_MAX_COLUMNS)!;
    expect(t.cells.map((row) => row.length)).toEqual([2, 1]);
    // No `a:tc` at the position: the cascade draws none there.
    expect(drawn(e)).toEqual(['A|B', 'C|·']);
  });

  test('a marked table draws the blank cell the save adds, and typing into it adds it in one undo step', async () => {
    const e = await open(bytes);
    expect(e.session!.store.replaceParts([markedPart(bytes)]).ok).toBe(true);
    const t = tableAsDrawn(frameOf(e), TABLE_MAX_COLUMNS)!;
    expect(t.cells[1]!.map((c) => c?.id === '')).toEqual([false, true]);
    expect(drawn(e)).toEqual(['A|B', 'C|']);
    expect(drawn(await reopened(e))).toEqual(drawn(e));
    caret(e, 1, 1);
    expect(e.input('text', 'x').ok).toBe(true);
    expect(drawn(e)).toEqual(['A|B', 'C|x']);
    expect(drawn(await reopened(e))).toEqual(['A|B', 'C|x']);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(children(children(find(rootOf(e), 'a:tbl'), 'a:tr')[1], 'a:tc')).toHaveLength(1);
    expect(drawn(e)).toEqual(['A|B', 'C|']);
  });
});
