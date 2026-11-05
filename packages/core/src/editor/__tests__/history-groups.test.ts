// Editor history groups: a typing burst, a drag, and a host handle are one undo step each, and
// undo and redo restore the selection before and after the step.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import type { EditorSelection } from '../../contracts/index.ts';
import { createPptxEditor, runToolbarCommand, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, name: 'T', x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('')) }) +
          shape({ id: 3, name: 'B', x: IN, y: 3 * IN, cx: IN, cy: IN }),
      },
    ],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const tree = (e: PptxEditorInstance): Element => {
  const part = e.session!.info.slides[0].partName;
  return find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!;
};
const byName = (e: PptxEditorInstance, name: string): Element =>
  elements(tree(e)).find((c) => attr(find(c, 'p:cNvPr'), 'name') === name)!;
const text = (e: PptxEditorInstance): string => textOf(find(byName(e, 'T'), 'p:txBody'));
const offX = (e: PptxEditorInstance): number => Number(attr(find(byName(e, 'B'), 'a:off'), 'x'));
const caret = (nodeId: string, offset: number): EditorSelection => ({
  kind: 'text',
  nodeId,
  anchor: { paragraph: 0, offset },
  focus: { paragraph: 0, offset },
});

/** Undo until nothing is left; returns the number of steps. */
function undoAll(e: PptxEditorInstance): number {
  let steps = 0;
  while (e.exec({ type: 'undo' }).ok) steps += 1;
  return steps;
}

describe('typing burst', () => {
  test('typed characters at a caret are one undo step; undo and redo restore the caret', async () => {
    const e = await open();
    const id = byName(e, 'T').id;
    e.setSelection(caret(id, 0));
    for (const ch of 'Hello') expect(e.input('text', ch).ok).toBe(true);
    expect(text(e)).toBe('Hello');
    expect(e.snapshot().selection).toEqual(caret(id, 5));

    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(text(e)).toBe('');
    expect(e.snapshot().selection).toEqual(caret(id, 0));
    expect(e.snapshot().canUndo).toBe(false);

    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(text(e)).toBe('Hello');
    expect(e.snapshot().selection).toEqual(caret(id, 5));
  });

  test('a caret move or a paragraph break between characters starts a new step', async () => {
    const e = await open();
    const id = byName(e, 'T').id;
    e.setSelection(caret(id, 0));
    e.input('text', 'a');
    e.input('text', 'b');
    e.setSelection(caret(id, 1));
    e.input('text', 'c');
    expect(text(e)).toBe('acb');
    expect(e.input('paragraph').ok).toBe(true);
    e.input('text', 'd');
    expect(undoAll(e)).toBe(4);
    expect(text(e)).toBe('');
  });
});

describe('drag and nudges', () => {
  test('the steps of one drag with one handle are one undo step; undo restores the shapes', async () => {
    const e = await open();
    const box = byName(e, 'B').id;
    const selected: EditorSelection = { kind: 'shapes', nodeIds: [box] };
    e.setSelection(selected);
    const drag = e.beginHistoryGroup();
    for (const x of [2, 3, 4]) {
      const items = [{ nodeId: box, x: x * IN, y: 3 * IN, cx: IN, cy: IN, rot: 0 }];
      expect(e.exec({ type: 'setTransforms', items }, { historyGroup: drag }).ok).toBe(true);
    }
    drag.end();
    expect(offX(e)).toBe(4 * IN);
    e.setSelection({ kind: 'none' });
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(offX(e)).toBe(IN);
    expect(e.snapshot().selection).toEqual(selected);
    expect(e.snapshot().canUndo).toBe(false);
    e.setSelection({ kind: 'none' });
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(offX(e)).toBe(4 * IN);
    expect(e.snapshot().selection).toEqual(selected);
  });

  test('consecutive nudges are one step; a new selection starts another', async () => {
    const e = await open();
    const box = byName(e, 'B').id;
    const text = byName(e, 'T').id;
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    for (let i = 0; i < 3; i += 1) e.exec({ type: 'nudge', dx: 9525, dy: 0 });
    e.setSelection({ kind: 'shapes', nodeIds: [text] });
    e.exec({ type: 'nudge', dx: 9525, dy: 0 });
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    e.exec({ type: 'nudge', dx: 9525, dy: 0 });
    expect(undoAll(e)).toBe(3);
  });
});

describe('host history group handles', () => {
  test('another handle, or none, closes the open step', async () => {
    const e = await open();
    const box = byName(e, 'B').id;
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    const first = e.beginHistoryGroup();
    const second = e.beginHistoryGroup();
    const nudge = { type: 'nudge', dx: 9525, dy: 0 } as const;
    e.exec(nudge, { historyGroup: first });
    e.exec(nudge, { historyGroup: first });
    e.exec(nudge, { historyGroup: second });
    expect(e.exec({ type: 'rotateObjects', degrees: 90 }).ok).toBe(true);
    e.exec(nudge, { historyGroup: first });
    expect(undoAll(e)).toBe(4);
  });

  test('a closed handle or a handle of another editor refuses and changes nothing', async () => {
    const e = await open();
    const other = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [byName(e, 'B').id] });
    const ended = e.beginHistoryGroup();
    ended.end();
    ended.end();
    expect(ended.state).toBe('closed');
    const nudge = { type: 'nudge', dx: 9525, dy: 0 } as const;
    const revision = e.snapshot().revision;
    expect(e.exec(nudge, { historyGroup: ended })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(e.can(nudge, { historyGroup: ended }).ok).toBe(false);
    const foreign = other.beginHistoryGroup();
    expect(e.exec(nudge, { historyGroup: foreign }).ok).toBe(false);
    expect(e.snapshot().revision).toBe(revision);
  });

  test('a selection change between two commands of one handle keeps one step', async () => {
    const e = await open();
    const box = byName(e, 'B').id;
    const label = byName(e, 'T').id;
    const group = e.beginHistoryGroup();
    const nudge = { type: 'nudge', dx: 9525, dy: 0 } as const;
    e.setSelection({ kind: 'shapes', nodeIds: [box] });
    expect(e.exec(nudge, { historyGroup: group }).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [label] });
    expect(e.exec(nudge, { historyGroup: group }).ok).toBe(true);
    group.end();
    expect(undoAll(e)).toBe(1);
    expect(offX(e)).toBe(IN);
  });

  test('value slots take the handle: the values of one slider drag are one step', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [byName(e, 'T').id] });
    const drag = e.beginHistoryGroup();
    for (const size of [20, 24, 28])
      expect(runToolbarCommand(e, 'font.size', size, { historyGroup: drag }).ok).toBe(true);
    drag.end();
    expect(runToolbarCommand(e, 'font.size', 32, { historyGroup: drag })).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(undoAll(e)).toBe(1);
  });

  test('a command that changes nothing keeps the selection that redo restores', async () => {
    const e = await open();
    const box = byName(e, 'B').id;
    const selected: EditorSelection = { kind: 'shapes', nodeIds: [box] };
    e.setSelection(selected);
    expect(e.exec({ type: 'nudge', dx: 9525, dy: 0 }).ok).toBe(true);
    e.setSelection({ kind: 'none' });
    const revision = e.snapshot().revision;
    expect(e.exec({ type: 'setNotes', index: 0, text: '' }).ok).toBe(true);
    expect(e.snapshot().revision).toBe(revision);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    e.setSelection({ kind: 'none' });
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(e.snapshot().selection).toEqual(selected);
  });

  test('loading a document closes every open handle', async () => {
    const e = await open();
    const handle = e.beginHistoryGroup();
    expect(handle.state).toBe('open');
    expect((await e.load(buildPptx({ slides: [{ shapes: '' }] }))).ok).toBe(true);
    expect(handle.state).toBe('closed');
  });
});
