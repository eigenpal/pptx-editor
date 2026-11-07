// The selection pane's rules, which both adapters share: the rows of the element tree, the tree
// keys, the selection a click makes, and the groups that open for a selected child.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  selectFromPane,
  selectionPaneAncestors,
  selectionPaneKey,
  selectionPaneRows,
  selectionPaneSelected,
  selectionPaneTopIds,
} from '../selection-pane.ts';
import type { ShapeSummary } from '../../contracts/index.ts';

const IN = 914_400;
const group =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Cards"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/>' +
  '<a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>' +
  shape({ id: 11, name: 'Card A', x: 0, y: 0, cx: IN, cy: IN }) +
  shape({ id: 12, name: 'Card B', x: IN, y: 0, cx: IN, cy: IN }) +
  '</p:grpSp>';

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, name: 'One', x: IN, y: IN, cx: IN, cy: IN }) +
          shape({ id: 3, name: 'Two', x: 2 * IN, y: IN, cx: IN, cy: IN }) +
          group,
      },
    ],
  });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

const names = (rows: readonly { name: string; level: number }[]): string[] =>
  rows.map((r) => `${r.level}:${r.name}`);
const byName = (shapes: readonly ShapeSummary[], name: string): ShapeSummary =>
  shapes.flatMap((s) => [s, ...(s.children ?? [])]).find((s) => s.name === name)!;

describe('selection pane rows', () => {
  test('the slide top first; a group starts closed and opens its children', async () => {
    const e = await open();
    const shapes = e.snapshot().shapes!;
    expect(names(selectionPaneRows(shapes, new Set()))).toEqual(['1:Cards', '1:Two', '1:One']);
    const cards = byName(shapes, 'Cards').nodeId;
    const rows = selectionPaneRows(shapes, new Set([cards]));
    expect(names(rows)).toEqual(['1:Cards', '2:Card B', '2:Card A', '1:Two', '1:One']);
    expect(rows[0]).toMatchObject({ branch: true, expanded: true, position: 1, siblings: 3 });
    expect(rows[1]).toMatchObject({ parent: cards, branch: false, position: 1, siblings: 2 });
    expect(selectionPaneTopIds(shapes)).toEqual(shapes.map((s) => s.nodeId));
  });

  test('the tree keys move, open, close, go to the parent, and rename', async () => {
    const e = await open();
    const shapes = e.snapshot().shapes!;
    const cards = byName(shapes, 'Cards').nodeId;
    const kid = byName(shapes, 'Card B').nodeId;
    const closed = selectionPaneRows(shapes, new Set());
    const opened = selectionPaneRows(shapes, new Set([cards]));
    expect(selectionPaneKey(closed, cards, 'ArrowRight')).toEqual({ expand: cards, open: true });
    expect(selectionPaneKey(opened, cards, 'ArrowRight')).toEqual({ focus: kid });
    expect(selectionPaneKey(opened, cards, 'ArrowLeft')).toEqual({ expand: cards, open: false });
    expect(selectionPaneKey(opened, kid, 'ArrowLeft')).toEqual({ focus: cards });
    expect(selectionPaneKey(opened, kid, 'ArrowUp')).toEqual({ focus: cards });
    expect(selectionPaneKey(opened, cards, 'End')).toEqual({ focus: opened.at(-1)!.nodeId });
    expect(selectionPaneKey(opened, kid, 'Home')).toEqual({ focus: cards });
    expect(selectionPaneKey(opened, kid, 'F2')).toEqual({ rename: kid });
    expect(selectionPaneKey(opened, kid, 'x')).toBeNull();
    expect(selectionPaneKey([], null, 'ArrowDown')).toBeNull();
  });
});

describe('selection from the pane', () => {
  test('a click selects; an extended click adds or removes in the same group only', async () => {
    const e = await open();
    const shapes = e.snapshot().shapes!;
    const cards = byName(shapes, 'Cards').nodeId;
    const rows = selectionPaneRows(shapes, new Set([cards]));
    const row = (name: string) => rows.find((r) => r.name === name)!;
    selectFromPane(e, row('One'), false);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [row('One').nodeId] });
    selectFromPane(e, row('Two'), true);
    expect(selectionPaneSelected(e.snapshot().selection)).toEqual(
      new Set([row('One').nodeId, row('Two').nodeId])
    );
    selectFromPane(e, row('One'), true);
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [row('Two').nodeId] });
    // A child of a group starts a new selection inside the group.
    selectFromPane(e, row('Card A'), true);
    expect(e.snapshot().selection).toEqual({
      kind: 'shapes',
      nodeIds: [row('Card A').nodeId],
      group: cards,
    });
    selectFromPane(e, row('Card A'), true);
    expect(e.snapshot().selection.kind).toBe('none');
  });

  test('a hidden element can be selected, and its group opens for a selected child', async () => {
    const e = await open();
    const shapes = e.snapshot().shapes!;
    const one = byName(shapes, 'One').nodeId;
    expect(e.exec({ type: 'setShapesVisible', nodeIds: [one], visible: false }).ok).toBe(true);
    const rows = selectionPaneRows(e.snapshot().shapes, new Set());
    selectFromPane(
      e,
      rows.find((r) => r.nodeId === one)!,
      false
    );
    expect(selectionPaneSelected(e.snapshot().selection)).toEqual(new Set([one]));
    const kid = byName(shapes, 'Card A').nodeId;
    expect(selectionPaneAncestors(shapes, new Set([kid]))).toEqual([
      byName(shapes, 'Cards').nodeId,
    ]);
    expect(selectionPaneAncestors(shapes, new Set([one]))).toEqual([]);
  });
});
