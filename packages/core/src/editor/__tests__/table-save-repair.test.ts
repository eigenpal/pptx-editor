// The save repair of merge markers (`repairTablesForSave` in `table-repair.ts`). A table with a
// structure mark (`table-edit-marks.ts`: a structure command changed it in a collaboration room)
// whose merge markup fails the grid saves with the markers the next table command writes. A
// merge and a concurrent delete of one of its covered rows leave such a table: the span reaches
// the row below. Every other table saves with the file's merge markup, after typing and after
// an undo too. No mark reaches a file, and the store does not change.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  canonicalPackageFingerprint,
  children,
  elements,
  find,
  readPptx,
  writePptx,
  editedTableGrid,
  resolveTableGrid,
  tableCellHasText,
  tableMergeFaults,
  tableSpanTrims,
  textOf,
  xmlPartOf,
  type Element,
  type PackagePart,
} from '../../store/index.ts';
import { packageViolations } from '../../store/store/part-checks.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { findMatches } from '../../automation/commands/find.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import {
  isStructureEdited,
  markStructureEdit,
  TABLE_EDIT_URI,
} from '../../automation/commands/table-edit-marks.ts';

const IN = 914_400;
const tr = (...cells: string[]): string => `<a:tr h="370840">${cells.join('')}</a:tr>`;

function deck(rows: string): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 3,
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: 370_840,
          columns: [IN, IN, IN],
          tblPr: '<a:tblPr firstRow="1" bandRow="1"/>',
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

/** Each cell, row by row: its text, then its merge attributes in brackets. */
const cells = (root: Element): string[] =>
  tcsOf(root).map((row) =>
    row
      .map((tc) => {
        const m = ['gridSpan', 'rowSpan', 'hMerge', 'vMerge']
          .filter((n) => attr(tc, n) !== undefined)
          .map((n) => `${n}=${attr(tc, n)}`)
          .join(',');
        return m ? `${textOf(tc)}[${m}]` : textOf(tc);
      })
      .join('|')
  );

/** Save and reopen; the reopened slide part's root. The store does not change. */
function saved(e: PptxEditorInstance): Element {
  const store = e.session!.store;
  const before = { revision: store.revision, print: canonicalPackageFingerprint(store.package) };
  const r = readPptx(e.save()!);
  expect({ revision: store.revision, print: canonicalPackageFingerprint(store.package) }).toEqual(
    before
  );
  if (!r.ok) throw new Error('reopen');
  return xmlPartOf(r.pkg, part(e))!.root;
}

/** The table-merge violations of a reopened slide. */
function mergeViolations(e: PptxEditorInstance): unknown[] {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return packageViolations(r.pkg).filter((v) => v.rule === 'table-merge');
}

const caretIn = (e: PptxEditorInstance, row: number, column: number, to = 0): void =>
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: to },
  });

// A merge of rows 0 and 1 of the first column whose row 1 a peer deleted: the span reaches "G".
const SPAN_OVER_G = tr(cell('AD', { attrs: ' rowSpan="2"' }), cell('B'), cell('C'));
const AFTER_DELETE = SPAN_OVER_G + tr(cell('G'), cell('H'), cell('I'));
const VALID = tr(cell('A'), cell('B'), cell('C')) + tr(cell('G'), cell('H'), cell('I'));
// Every fault a file can hold: an orphan, an overlap, a span past the grid, an unmarked cell.
const FAULTY =
  tr(cell('A'), cell('B', { attrs: ' rowSpan="2"' }), cell('C', { attrs: ' vMerge="1"' })) +
  tr(cell('D', { attrs: ' gridSpan="2"' }), cell('E'), cell('F', { attrs: ' rowSpan="3"' }));

/** The slide part of `bytes`; `marked` gives its table a structure mark. */
function slidePart(bytes: Uint8Array, marked: boolean): PackagePart {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('read');
  const name = [...r.pkg.parts.keys()].find((n) => /slides\/slide1\.xml$/.test(n))!;
  const batch = new OpBatch(r.pkg);
  const tree = find(xmlPartOf(r.pkg, name)!.root, 'p:spTree')!;
  const frame = elements(tree).find((c) => c.local === 'graphicFrame')!;
  if (marked) markStructureEdit(batch, name, frame.id);
  return batch.pkg.parts.get(name)!;
}

/**
 * Put the slide part of `bytes` into `e` as a peer's part (`replaceParts`, as a room does). By
 * default its table has a structure mark, as a peer's structure command writes it.
 */
function receive(e: PptxEditorInstance, bytes: Uint8Array, marked = true): void {
  expect(e.session!.store.replaceParts([slidePart(bytes, marked)]).ok).toBe(true);
}

/** Whether the store's table has a structure mark. */
const markedInStore = (e: PptxEditorInstance): boolean =>
  isStructureEdited(elements(find(slideRoot(e), 'p:spTree')!).find((c) => c.id === frameId(e)));

describe('the save repair of merge markers', () => {
  test('a span that a peer left over the next row ends before it, and that row shows', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(AFTER_DELETE));
    expect(cells(slideRoot(e))).toEqual(['AD[rowSpan=2]|B|C', 'G|H|I']);
    expect(cells(saved(e))).toEqual(['AD|B|C', 'G|H|I']);
    expect(mergeViolations(e)).toEqual([]);
  });

  test('the saved markup is what the next table command writes', async () => {
    for (const rows of [AFTER_DELETE, SPAN_OVER_G + tr(cell(''), cell('H'), cell('I'))]) {
      const e = await open(deck(VALID));
      receive(e, deck(rows));
      const repaired = cells(saved(e));
      caretIn(e, 0, 1, 1);
      expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(cells(slideRoot(e))).toEqual(repaired);
      expect(cells(saved(e))).toEqual(repaired);
    }
  });

  test('a blank cell under the span stays covered and gets vMerge', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(SPAN_OVER_G + tr(cell(''), cell('H'), cell('I'))));
    expect(cells(saved(e))).toEqual(['AD[rowSpan=2]|B|C', '[vMerge=1]|H|I']);
    expect(mergeViolations(e)).toEqual([]);
  });

  test('an unedited table with faults saves as the file wrote it', async () => {
    for (const rows of [AFTER_DELETE, FAULTY]) {
      const e = await open(deck(rows));
      expect(mergeViolations(e).length).toBeGreaterThan(0);
      expect(cells(saved(e))).toEqual(cells(slideRoot(e)));
    }
  });

  test('typing in a table with faults keeps its merge markup', async () => {
    const e = await open(deck(FAULTY));
    const before = cells(slideRoot(e));
    caretIn(e, 0, 0);
    expect(e.input('text', 'x').ok).toBe(true);
    expect(cells(saved(e))).toEqual([before[0].replace('A|', 'xA|'), before[1]]);
  });

  test('a peer part without a structure mark keeps its faults; one with the mark is repaired', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(AFTER_DELETE), false);
    expect(cells(saved(e))).toEqual(['AD[rowSpan=2]|B|C', 'G|H|I']);
    receive(e, deck(FAULTY));
    expect(mergeViolations(e)).toEqual([]);
  });

  test('after an undo of a table command, the table saves as the file wrote it again', async () => {
    const e = await open(deck(AFTER_DELETE));
    const file = cells(slideRoot(e));
    caretIn(e, 1, 2);
    expect(e.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
    expect(mergeViolations(e)).toEqual([]);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(cells(saved(e))).toEqual(file);
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(mergeViolations(e)).toEqual([]);
  });

  test('a marked table in mc:Fallback loses its mark and gets its merge markers repaired', async () => {
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const frame = table({
      id: 3,
      x: IN,
      y: IN,
      cx: 3 * IN,
      cy: 370_840,
      columns: [IN, IN, IN],
      tblPr: '<a:tblPr/>',
      rows: AFTER_DELETE,
    });
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="p14" ` +
            `xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${frame}` +
            `</mc:Choice><mc:Fallback>${frame}</mc:Fallback></mc:AlternateContent>`,
        },
      ],
    });
    const e = await open(bytes);
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('read');
    const batch = new OpBatch(r.pkg);
    const fallback = find(xmlPartOf(r.pkg, part(e))!.root, 'mc:Fallback')!;
    const marked = elements(fallback).find((c) => c.local === 'graphicFrame')!;
    markStructureEdit(batch, part(e), marked.id);
    expect(e.session!.store.replaceParts([batch.pkg.parts.get(part(e))!]).ok).toBe(true);
    const bytesOut = e.save()!;
    expect(new TextDecoder().decode(bytesOut)).not.toContain(TABLE_EDIT_URI);
    const out = readPptx(bytesOut);
    if (!out.ok) throw new Error('reopen');
    const root = xmlPartOf(out.pkg, part(e))!.root;
    expect(cells(find(root, 'mc:Fallback')!)).toEqual(['AD|B|C', 'G|H|I']);
    // The unmarked frame in mc:Choice keeps the file's markup.
    expect(cells(find(root, 'mc:Choice')!)).toEqual(['AD[rowSpan=2]|B|C', 'G|H|I']);
  });

  test('a mark never reaches a file, and a table command outside a room writes none', async () => {
    const e = await open(deck(VALID));
    caretIn(e, 0, 0);
    expect(e.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    expect(markedInStore(e)).toBe(false);
    receive(e, deck(AFTER_DELETE));
    expect(markedInStore(e)).toBe(true);
    expect(new TextDecoder().decode(e.save()!)).not.toContain(TABLE_EDIT_URI);
  });

  test('a mark in a file is forged: the open removes it, and the table keeps its markup', async () => {
    const r = readPptx(deck(AFTER_DELETE));
    if (!r.ok) throw new Error('read');
    const slide = slidePart(deck(AFTER_DELETE), true);
    const parts = new Map(r.pkg.parts);
    parts.set(slide.partName, slide);
    const e = await open(writePptx({ parts }));
    expect(markedInStore(e)).toBe(false);
    expect(cells(saved(e))).toEqual(['AD[rowSpan=2]|B|C', 'G|H|I']);
  });

  test('a copy of a marked table carries no mark, and its paste holds none', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(VALID));
    e.setSelection({ kind: 'shapes', nodeIds: [frameId(e)] });
    const payload = e.copy()!;
    expect(JSON.stringify(payload)).not.toContain(TABLE_EDIT_URI);
    expect(e.exec({ type: 'paste', payload }).ok).toBe(true);
    const frames = elements(find(slideRoot(e), 'p:spTree')!).filter(
      (c) => c.local === 'graphicFrame'
    );
    expect(frames.map((f) => isStructureEdited(f))).toEqual([true, false]);
  });
});

/** The table `e` draws: per row, each drawn cell as `text[rows x columns]`, `·` when covered. */
function drawn(e: PptxEditorInstance): string[] {
  const item = e.session!.layout(0)!.items.find((i) => i.element.kind === 'graphicFrame')!;
  const table = item.element.kind === 'graphicFrame' ? item.element.table! : undefined;
  const byId = new Map(tcsOf(slideRoot(e)).flatMap((row) => row.map((tc) => [tc.id, tc])));
  const out: string[] = [];
  for (const c of table!.cells) {
    if (c.col === 0) out.push('');
    const spans = c.gridSpan > 1 || c.rowSpan > 1 ? `[${c.rowSpan}x${c.gridSpan}]` : '';
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : textOf(byId.get(c.nodeId)) + spans}`;
  }
  return out;
}

describe('the canvas of a marked table', () => {
  const BLANK_BELOW = SPAN_OVER_G + tr(cell(''), cell('H'), cell('I'));

  test('a marked table draws the grid its save writes, with no store write', async () => {
    for (const rows of [AFTER_DELETE, BLANK_BELOW, FAULTY]) {
      const e = await open(deck(VALID));
      receive(e, deck(rows));
      const revision = e.session!.store.revision;
      const canvas = drawn(e);
      expect(await open(e.save()!).then(drawn)).toEqual(canvas);
      expect(e.session!.store.revision).toBe(revision);
    }
  });

  test('the row below a span that a peer left shows; a blank one stays covered', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(AFTER_DELETE));
    expect(drawn(e)).toEqual(['AD|B|C', 'G|H|I']);
    receive(e, deck(BLANK_BELOW));
    expect(drawn(e)).toEqual(['AD[2x1]|B|C', '·|H|I']);
  });

  test('a table without a mark draws its file markup, faults included', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(AFTER_DELETE), false);
    expect(drawn(e)).toEqual(['AD[2x1]|B|C', '·|H|I']);
    const file = await open(deck(FAULTY));
    expect(drawn(file)).toEqual(await open(file.save()!).then(drawn));
  });

  test('a marked table wider than the repair reads draws and saves its markup', async () => {
    const wide = (rows: string): Uint8Array =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: table({
              id: 3,
              x: 0,
              y: 0,
              cx: 76 * 100_000,
              cy: 370_840,
              columns: Array.from({ length: 76 }, () => 100_000),
              tblPr: '<a:tblPr/>',
              rows,
            }),
          },
        ],
      });
    const rest = Array.from({ length: 75 }, () => cell('x')).join('');
    const e = await open(wide(tr(cell('A'), rest) + tr(cell('G'), rest)));
    receive(e, wide(tr(cell('A', { attrs: ' rowSpan="2"' }), rest) + tr(cell('G'), rest)));
    expect(drawn(e)[1]!.startsWith('·|')).toBe(true);
    expect(await open(e.save()!).then(drawn)).toEqual(drawn(e));
  });

  test('find reads the cells the canvas draws', async () => {
    const e = await open(deck(VALID));
    receive(e, deck(AFTER_DELETE), false);
    expect(findMatches(e.session!.store.package, 'G')).toHaveLength(0);
    receive(e, deck(AFTER_DELETE));
    expect(findMatches(e.session!.store.package, 'G')).toHaveLength(1);
  });
});

describe('the merge markup helpers', () => {
  test('editedTableGrid is resolveTableGrid for a table without faults, and trims once', async () => {
    const valid = tr(cell('AD', { attrs: ' rowSpan="2"' }), cell('B'), cell('C'));
    for (const rows of [
      VALID,
      valid + tr(cell('', { attrs: ' vMerge="1"' }), cell('E'), cell('F')),
    ]) {
      const tcs = tcsOf(slideRoot(await open(deck(rows))));
      expect(tableSpanTrims(tcs, tableCellHasText)).toEqual([]);
      expect(editedTableGrid(tcs, 3, tableCellHasText)).toEqual(resolveTableGrid(tcs, 3));
    }
    const tcs = tcsOf(slideRoot(await open(deck(AFTER_DELETE))));
    expect(tableSpanTrims(tcs, tableCellHasText)).toEqual([
      { row: 0, col: 0, name: 'rowSpan', value: 1 },
    ]);
    expect(editedTableGrid(tcs, 3, tableCellHasText).blocks).toHaveLength(6);
  });

  test('a frame keeps at most 64 marks: a step past them replaces the oldest', async () => {
    const e = await open(deck(VALID));
    const batch = new OpBatch(e.session!.store.package);
    for (let i = 0; i < 70; i += 1) markStructureEdit(batch, part(e), frameId(e));
    const frame = batch.element(part(e), frameId(e));
    const cNvPr = find(find(frame, 'p:nvGraphicFramePr')!, 'p:cNvPr')!;
    expect(elements(cNvPr).filter((c) => attr(c, 'uri') === TABLE_EDIT_URI)).toHaveLength(64);
  });

  test('tableMergeFaults names each fault of a table', async () => {
    const e = await open(deck(FAULTY));
    const tcs = tcsOf(slideRoot(e));
    const faults = tableMergeFaults(tcs, resolveTableGrid(tcs, 3));
    expect(faults.map((f) => `${f.row},${f.col} ${f.fault}`)).toEqual([
      '1,0 overlap',
      '1,2 past-grid',
      '0,2 orphan',
      '1,1 unmarked',
    ]);
    // Positions without an `a:tc` (short rows) have no fault.
    const short = [[tcs[0][0]], []];
    expect(tableMergeFaults(short, resolveTableGrid(short, 3))).toEqual([]);
  });
});
