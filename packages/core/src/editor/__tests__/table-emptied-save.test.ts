// A table that concurrent deletes left with no grid column or no row (`isEmptiedTable`). Two
// peers who delete the last columns of a table at the same time leave an empty `a:tblGrid`,
// and a row that one of them inserted can keep a cell of a column the other deleted, which
// fails the `table-grid` part check (ECMA-376 Part 1, 21.1.3.16 tc: one `a:tc` per
// `a:gridCol`). A local delete of the last column deletes the whole table, so a table with a
// structure mark and no column or row is not drawn or listed in the selection pane, and the
// save leaves its frame out, with its animations and connector ends. The room keeps them, so a
// column that comes back brings them back. A table without a mark saves as the file wrote it.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../../../scripts/fixtures/timing.ts';
import {
  canonicalPackageFingerprint,
  elements,
  find,
  isEmptiedTable,
  readPptx,
  xmlPartOf,
  type Element,
  type PackagePart,
} from '../../store/index.ts';
import { packageViolations } from '../../store/store/part-checks.ts';
import { findMatches } from '../../automation/commands/find.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import { markStructureEdit } from '../../automation/commands/table-edit-marks.ts';
import { asSavedForCopy, repairTablesForSave } from '../../automation/commands/table-repair.ts';

const IN = 914_400;
const tr = (...cells: string[]): string => `<a:tr h="370840">${cells.join('')}</a:tr>`;
const VALID = tr(cell('A'), cell('B')) + tr(cell('C'), cell('D'));

/** A connector (id 4) glued from the box (id 2) to the table (id 3). */
const CONNECTOR =
  '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="Connector 3"/><p:cNvCxnSpPr>' +
  '<a:stCxn id="2" idx="3"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>' +
  '<p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="0" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>';

/**
 * A slide with a box (id 2) and a table (id 3) that `rows` and `columns` fill. A `linked` slide
 * also animates both and glues a connector (id 4) from the box to the table.
 */
function deck(rows: string, columns: readonly number[], linked = false): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, x: 0, y: 0, cx: IN, cy: IN, text: para(run('Box')) }) +
          table({
            id: 3,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: 741_680,
            columns,
            tblPr: '<a:tblPr/>',
            rows,
          }) +
          (linked ? CONNECTOR : ''),
        ...(linked ? { ext: clickEntrances([2, 3]) } : {}),
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
const tree = (root: Element): Element => find(root, 'p:spTree')!;
const frameOf = (root: Element): Element | undefined =>
  elements(tree(root)).find((c) => c.local === 'graphicFrame');
const store = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, part(e))!.root;

/**
 * Put the slide part of `bytes` into `e` as a peer's part (`replaceParts`, as a room does),
 * with a structure mark on its table unless `marked` is false.
 */
function receive(e: PptxEditorInstance, bytes: Uint8Array, marked = true): void {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('read');
  const name = part(e);
  const batch = new OpBatch(r.pkg);
  if (marked) markStructureEdit(batch, name, frameOf(xmlPartOf(r.pkg, name)!.root)!.id);
  const slide: PackagePart = batch.pkg.parts.get(name)!;
  expect(e.session!.store.replaceParts([slide]).ok).toBe(true);
}

/** Save and reopen with no store write; the reopened package. */
function saved(e: PptxEditorInstance) {
  const s = e.session!.store;
  const before = { revision: s.revision, print: canonicalPackageFingerprint(s.package) };
  const r = readPptx(e.save()!);
  expect({ revision: s.revision, print: canonicalPackageFingerprint(s.package) }).toEqual(before);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
}

/** An element's markup without node ids, to compare a saved element with the store's. */
const markup = (el: Element | undefined): string =>
  JSON.stringify(el, (key, value: unknown) => (key === 'id' ? undefined : value));

/** The kinds of the items the canvas draws. */
const drawnKinds = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.kind);

// No grid column; the second row keeps the cell of a column that a peer deleted.
const NO_COLUMN = tr() + tr(cell('X'));

describe('a marked table with no grid column or no row', () => {
  for (const [name, rows, columns] of [
    ['no grid column, with a cell left', NO_COLUMN, []],
    ['no grid column and no cell', tr() + tr(), []],
    ['no row', '', [IN, IN]],
  ] as const)
    test(`${name}: the save leaves the frame out, and the canvas does not draw it`, async () => {
      const e = await open(deck(VALID, [IN, IN]));
      receive(e, deck(rows, columns));
      expect(isEmptiedTable(frameOf(store(e)))).toBe(true);
      expect(drawnKinds(e)).toEqual(['shape']);
      const pkg = saved(e);
      expect(packageViolations(pkg)).toEqual([]);
      expect(frameOf(xmlPartOf(pkg, part(e))!.root)).toBeUndefined();
      // The store keeps the room's state.
      expect(frameOf(store(e))).toBeDefined();
    });

  test('the frame goes with its animations and connector end; a column back brings them back', async () => {
    const e = await open(deck(VALID, [IN, IN], true));
    /** The `spid` of each `local` in the saved slide, and the ids its connector is glued to. */
    const links = (root: Element) => {
      const ids = (local: string, name: string) => {
        const out: string[] = [];
        const walk = (el: Element) => {
          for (const c of elements(el)) {
            if (c.local === local) out.push(c.attrs.find((a) => a.local === name)!.value);
            walk(c);
          }
        };
        walk(root);
        return out;
      };
      return {
        targets: ids('spTgt', 'spid'),
        builds: ids('bldP', 'spid'),
        glued: [...ids('stCxn', 'id'), ...ids('endCxn', 'id')],
      };
    };
    receive(e, deck(NO_COLUMN, [], true));
    expect(links(xmlPartOf(saved(e), part(e))!.root)).toEqual({
      targets: ['2'],
      builds: ['2'],
      glued: ['2'],
    });
    // The room kept them: when a column comes back (an undo), the save writes them again.
    receive(e, deck(tr(cell('A')), [IN], true));
    const back = saved(e);
    expect(packageViolations(back)).toEqual([]);
    expect(links(xmlPartOf(back, part(e))!.root)).toEqual({
      targets: ['2', '3'],
      builds: ['2', '3'],
      glued: ['2', '3'],
    });
  });

  test('the selection pane does not list it, and find does not reach its text', async () => {
    const e = await open(deck(VALID, [IN, IN]));
    receive(e, deck(tr() + tr(cell('Kept')), []));
    expect(e.snapshot().shapes!.map((s) => s.kind)).toEqual(['shape']);
    expect(findMatches(e.session!.store.package, 'Kept')).toEqual([]);
  });

  test('a table without a mark saves as the file wrote it', async () => {
    const e = await open(deck(VALID, [IN, IN]));
    receive(e, deck(NO_COLUMN, []), false);
    expect(isEmptiedTable(frameOf(store(e)))).toBe(false);
    expect(markup(frameOf(xmlPartOf(saved(e), part(e))!.root))).toBe(markup(frameOf(store(e))));
    const file = await open(deck('', [IN, IN]));
    expect(markup(frameOf(xmlPartOf(saved(file), part(file))!.root))).toBe(
      markup(frameOf(store(file)))
    );
  });

  test('a marked table with a column and a row is not emptied', async () => {
    const e = await open(deck(VALID, [IN, IN]));
    receive(e, deck(tr(cell('A')), [IN]));
    expect(isEmptiedTable(frameOf(store(e)))).toBe(false);
    expect(drawnKinds(e)).toEqual(['shape', 'graphicFrame']);
  });

  test('a copy leaves the table out, as the save does', async () => {
    const e = await open(deck(VALID, [IN, IN]));
    receive(e, deck(NO_COLUMN, []));
    const frame = frameOf(store(e))!;
    expect(asSavedForCopy(e.session!.store.package, part(e), frame)).toBeUndefined();
    e.setSelection({ kind: 'shapes', nodeIds: [frame.id] });
    expect(e.copy()).toBeNull();
  });

  test('in a group, the copy and the save keep the group and leave the table out', () => {
    const grouped = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>' +
            shape({ id: 2, x: 0, y: 0, cx: IN, cy: IN, text: para(run('Box')) }) +
            table({
              id: 3,
              x: IN,
              y: 0,
              cx: IN,
              cy: IN,
              columns: [],
              tblPr: '<a:tblPr/>',
              rows: NO_COLUMN,
            }) +
            '</p:grpSp>',
        },
      ],
    });
    const r = readPptx(grouped);
    if (!r.ok) throw new Error('read');
    const name = [...r.pkg.parts.keys()].find((n) => /slides\/slide1\.xml$/.test(n))!;
    const group = () =>
      elements(tree(xmlPartOf(batch.pkg, name)!.root)).find((c) => c.local === 'grpSp')!;
    const batch = new OpBatch(r.pkg);
    markStructureEdit(batch, name, elements(group()).find((c) => c.local === 'graphicFrame')!.id);
    const kinds = (el: Element | undefined) => elements(el).map((c) => c.local);
    expect(kinds(group())).toEqual(['nvGrpSpPr', 'grpSpPr', 'sp', 'graphicFrame']);
    expect(kinds(asSavedForCopy(batch.pkg, name, group()))).toEqual(['nvGrpSpPr', 'grpSpPr', 'sp']);
    const saved = repairTablesForSave(batch.pkg);
    expect(
      kinds(elements(tree(xmlPartOf(saved, name)!.root)).find((c) => c.local === 'grpSp'))
    ).toEqual(['nvGrpSpPr', 'grpSpPr', 'sp']);
    expect(packageViolations(saved)).toEqual([]);
  });
});
