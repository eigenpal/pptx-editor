// The commit checks of touched parts (`store/store/part-checks.ts`) against real editing: every
// synthetic fixture opens, takes the common commands on every slide, and saves, and no command
// is refused for a broken part. A refused remote transaction comes back to the collaboration
// session, never dropped.

import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, cell, shape, table } from '../../../../../scripts/fixtures/builder.ts';
import { children, find, readPptx, xmlPartOf, type Element } from '../../store/index.ts';
import type { CommandResult, EditorCommand } from '../../contracts/index.ts';
import { collaborationPort } from '../modules.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');
const decks = readdirSync(FIXTURES).filter((f) => f.endsWith('.pptx'));

const MEMBERS = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp']);
const INVARIANT = new Set(['tree-invariant', 'package-invariant']);

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const slideRoot = (e: PptxEditorInstance, i: number): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[i].partName)!.root;

/** The top-level shapes of slide `i`. */
const members = (e: PptxEditorInstance, i: number): Element[] =>
  (find(slideRoot(e, i), 'p:spTree')?.children ?? []).filter(
    (c): c is Element => c.type === 'element' && MEMBERS.has(c.kind)
  );

/** A command may be refused for its own reasons, never for a part the store found broken. */
function notBroken(deck: string, command: EditorCommand | string, r: CommandResult): void {
  if (!r.ok && INVARIANT.has((r.detail ?? '').split(' ')[0]))
    throw new Error(`${deck}: ${JSON.stringify(command)} left a broken part (${r.detail})`);
}

function run(deck: string, e: PptxEditorInstance, command: EditorCommand): CommandResult {
  const r = e.exec(command);
  notBroken(deck, command, r);
  return r;
}

function editSlide(deck: string, e: PptxEditorInstance, i: number): void {
  e.goToSlide(i);
  const shapes = members(e, i);
  // Text: type into the first text body.
  const text = shapes.find((s) => s.kind === 'sp' && find(s, 'p:txBody'));
  if (text) {
    const at = { paragraph: 0, offset: 0 };
    e.setSelection({ kind: 'text', nodeId: text.id, anchor: at, focus: at });
    notBroken(deck, 'input', e.input('text', 'x'));
    notBroken(deck, 'paragraph', e.input('paragraph'));
  }
  // Tables: insert and delete a row and a column, merge two cells, and split them again.
  for (const frame of shapes.filter((s) => find(s, 'a:tbl'))) {
    const cellAt = { row: 0, column: 0 };
    e.setSelection({ kind: 'cells', nodeId: frame.id, anchor: cellAt, focus: cellAt });
    for (const command of [
      { type: 'insertRow', where: 'below' },
      { type: 'insertColumn', where: 'right' },
      { type: 'deleteRow' },
      { type: 'deleteColumn' },
    ] as const)
      run(deck, e, command);
    const cols = children(find(find(slideRoot(e, i), 'a:tbl'), 'a:tblGrid'), 'a:gridCol').length;
    if (cols > 1 && members(e, i).some((m) => m.id === frame.id)) {
      e.setSelection({
        kind: 'cells',
        nodeId: frame.id,
        anchor: cellAt,
        focus: { row: 0, column: 1 },
      });
      if (run(deck, e, { type: 'mergeCells' }).ok) {
        e.setSelection({ kind: 'cells', nodeId: frame.id, anchor: cellAt, focus: cellAt });
        run(deck, e, { type: 'splitCell', rows: 1, cols: 2 });
      }
    }
  }
  // Shapes: duplicate, paste, nudge, group, ungroup, and delete the copies.
  const ids = members(e, i).map((s) => s.id);
  if (ids.length > 0) {
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    const payload = e.copy();
    run(deck, e, { type: 'duplicate' });
    run(deck, e, { type: 'nudge', dx: 12_700, dy: 0 });
    if (ids.length > 1) {
      run(deck, e, { type: 'group' });
      run(deck, e, { type: 'ungroup' });
    }
    run(deck, e, { type: 'delete' });
    if (payload) {
      run(deck, e, { type: 'paste', payload });
      run(deck, e, { type: 'delete' });
    }
  }
  e.setSelection({ kind: 'none' });
  run(deck, e, { type: 'insertShape', preset: 'ellipse' });
  run(deck, e, { type: 'insertTextBox' });
  run(deck, e, { type: 'insertTable', rows: 2, cols: 2 });
}

describe('every fixture edits and saves without a broken part', () => {
  for (const deck of decks)
    test(deck, async () => {
      const e = await open(new Uint8Array(readFileSync(join(FIXTURES, deck))));
      const count = e.session!.info.slides.length;
      for (let i = 0; i < count; i += 1) editSlide(deck, e, i);
      run(deck, e, { type: 'addSlide' });
      run(deck, e, { type: 'duplicateSlide', index: 0 });
      run(deck, e, { type: 'moveSlide', from: 0, to: e.session!.info.slides.length - 1 });
      run(deck, e, { type: 'deleteSlide', index: 0 });
      // Undo and redo restore checked revisions.
      run(deck, e, { type: 'undo' });
      run(deck, e, { type: 'redo' });
      const saved = e.save();
      expect(saved).toBeTruthy();
      const reopened = readPptx(saved!);
      expect(reopened.ok).toBe(true);
    });
});

describe('a command the store refuses says why', () => {
  test('deleting a slide that another slide still names by an unknown element', async () => {
    // Slide 2 names slide 1 through `p:tags`, which `deleteSlide` cannot drop. Without the
    // check the delete committed and saved a relationship to the deleted slide part.
    const e = await open(
      buildPptx({
        slides: [
          { shapes: '' },
          {
            shapes: shape({
              id: 2,
              x: 0,
              y: 0,
              cx: 100,
              cy: 100,
              ph: '<p:custDataLst><p:tags r:id="rId9"/></p:custDataLst>',
            }),
            rels: [
              [
                'rId9',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                'slide1.xml',
              ],
            ],
          },
        ],
      })
    );
    const revision = e.session!.store.revision;
    const r = e.exec({ type: 'deleteSlide', index: 0 });
    expect(r).toEqual({
      ok: false,
      reason: 'errors.unsupported',
      detail:
        'package-invariant relationship-target /ppt/slides/_rels/slide2.xml.rels: ' +
        'relationship rId9 targets /ppt/slides/slide1.xml, which the transaction removed',
    });
    expect(e.session!.store.revision).toBe(revision);
    expect(e.session!.info.slides.length).toBe(2);
  });
});

describe('collaboration', () => {
  test('a remote transaction that breaks a table comes back to the session', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: table({
              id: 2,
              x: 0,
              y: 0,
              cx: 2_000_000,
              cy: 741_680,
              columns: [1_000_000, 1_000_000],
              tblPr: '<a:tblPr/>',
              rows: `<a:tr h="370840">${cell('a')}${cell('b')}</a:tr>`,
            }),
          },
        ],
      })
    );
    const store = e.session!.store;
    const part = e.session!.info.slides[0].partName;
    const tc = find(xmlPartOf(store.package, part)!.root, 'a:tc')!;
    const revision = store.revision;
    const r = collaborationPort(store).applyRemote([{ op: 'removeNode', part, node: tc.id }]);
    expect(r.ok ? null : r.rejection).toMatchObject({
      code: 'tree-invariant',
      rule: 'table-grid',
      part,
    });
    expect(store.revision).toBe(revision);
  });
});
