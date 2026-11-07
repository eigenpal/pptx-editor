// `can` answers within a frame on large targets: an arrange command on more than 250 shapes, a
// Duplicate of shapes with more than 20,000 XML nodes, and Apply Layout or Reset Slide on a slide
// of more than 20,000 XML nodes check only the editing mode. Other commands on more than 250
// shapes check the first 250 first, and a refusal there checks them all. A text format check reads at most 64 runs, but the
// command formats every run. Formatting answers hold through a typing burst.

import { expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { LARGE_SELECTION, TYPING_SETTLE_MS } from '../can-cache.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

const shapes = (n: number) =>
  Array.from({ length: n }, (_, i) =>
    shape({
      id: i + 2,
      x: (i % 50) * IN * 0.2,
      y: Math.floor(i / 50) * IN * 0.2,
      cx: IN / 10,
      cy: IN / 10,
    })
  ).join('');

/** One text box of `runs` runs: about four XML nodes each. */
const textBox = (runs: number) =>
  shape({
    id: 2,
    x: IN,
    y: IN,
    cx: 6 * IN,
    cy: 3 * IN,
    text: para(Array.from({ length: runs }, (_, i) => run(`w${i} `)).join('')),
  });

async function open(slides: string[], mode?: 'view'): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    ...(mode ? { mode } : {}),
  });
  const r = await editor.load(buildPptx({ slides: slides.map((s) => ({ shapes: s })) }));
  expect(r.ok).toBe(true);
  return editor;
}

const ids = (editor: PptxEditorInstance, slide = 0) =>
  editor.session!.layout(slide)!.items.map((i) => i.element.nodeId);

test('more than 250 shapes check only the mode; 250 shapes check in full', async () => {
  const editor = await open([shapes(LARGE_SELECTION + 1)]);
  const all = ids(editor);
  editor.setSelection({ kind: 'shapes', nodeIds: all.slice(0, LARGE_SELECTION) });
  // Plain shapes are no group: the full check refuses.
  expect(editor.can({ type: 'ungroup' }).ok).toBe(false);
  editor.setSelection({ kind: 'shapes', nodeIds: all });
  expect(editor.can({ type: 'ungroup' })).toEqual({ ok: true });
  // The command itself checks in full and refuses.
  expect(editor.exec({ type: 'ungroup' }).ok).toBe(false);
  editor.destroy();
  const viewer = await open([shapes(LARGE_SELECTION + 1)], 'view');
  viewer.setSelection({ kind: 'shapes', nodeIds: ids(viewer) });
  expect(viewer.can({ type: 'duplicate' })).toEqual({ ok: false, reason: 'disabled.viewing' });
  viewer.destroy();
});

test('one shape of many nodes takes the quick answer for Duplicate only', async () => {
  const editor = await open([textBox(7_000)]);
  const [box] = ids(editor);
  editor.setSelection({ kind: 'shapes', nodeIds: [box!] });
  expect(editor.can({ type: 'duplicate' })).toEqual({ ok: true });
  // Group needs two shapes: the full check still says so.
  expect(editor.can({ type: 'group' }).ok).toBe(false);
  editor.destroy();
});

test('Apply Layout checks only the mode on a slide of many nodes', async () => {
  const editor = await open([textBox(7_000), shapes(3)]);
  const unknown = { type: 'applyLayout', layoutId: 'no-such-layout', indices: [0] } as const;
  editor.setSelection({ kind: 'none' });
  expect(editor.can(unknown)).toEqual({ ok: true });
  expect(editor.exec(unknown).ok).toBe(false);
  // The slides the command names count, not the current one.
  expect(editor.can({ ...unknown, indices: [1] }).ok).toBe(false);
  editor.selectSlide(1);
  expect(editor.can(unknown)).toEqual({ ok: true });
  editor.destroy();
});

test('a text format check of a long text box is quick, and the command formats every run', async () => {
  const editor = await open([textBox(5_000)]);
  const [box] = ids(editor);
  editor.setSelection({ kind: 'shapes', nodeIds: [box!] });
  expect(editor.can({ type: 'toggleBold' })).toEqual({ ok: true });
  expect(editor.exec({ type: 'toggleBold' }).ok).toBe(true);
  const item = editor.session!.layout(0)!.items[0]!;
  const runs = item.kind === 'shape' ? (item.text?.lines.flatMap((l) => l.runs) ?? []) : [];
  expect(runs.length).toBeGreaterThan(100);
  expect(runs.every((r) => r.props.bold)).toBe(true);
  editor.destroy();
});

test('formatting answers hold through a typing burst, and are exact again after it', async () => {
  const editor = await open([textBox(3)]);
  const [box] = ids(editor);
  const at = { paragraph: 0, offset: 0 };
  editor.setSelection({ kind: 'text', nodeId: box!, anchor: at, focus: at });
  editor.input('text', 'a');
  const held = editor.can({ type: 'toggleBold' });
  editor.input('text', 'b');
  // The same answer, not a new check, while keys follow each other.
  expect(editor.can({ type: 'toggleBold' })).toBe(held);
  await new Promise((r) => setTimeout(r, TYPING_SETTLE_MS + 100));
  const settled = editor.can({ type: 'toggleBold' });
  expect(settled).toEqual(held);
  expect(settled).not.toBe(held);
  editor.destroy();
});

test('a large selection keeps the refusals of slide commands and of arguments', async () => {
  const editor = await open([shapes(LARGE_SELECTION + 1)]);
  editor.setSelection({ kind: 'shapes', nodeIds: ids(editor) });
  // The only slide: Delete Slide refuses, as with no selection.
  expect(editor.can({ type: 'deleteSlide', index: 0 })).toEqual({
    ok: false,
    reason: 'disabled.lastSlide',
  });
  expect(editor.can({ type: 'deleteSlide', index: 99 }).ok).toBe(false);
  expect(editor.can({ type: 'setLink', url: 'javascript:alert(1)' })).toEqual({
    ok: false,
    reason: 'errors.unsafeLink',
  });
  expect(editor.can({ type: 'setFontSize', points: 99_999 }).ok).toBe(false);
  // A valid argument passes on the first shapes.
  expect(editor.can({ type: 'setShapeFill', color: '#FF0000' })).toEqual({ ok: true });
  // Refusals that read the editor state or the text stay exact.
  expect(editor.can({ type: 'pasteFormat' })).toEqual({
    ok: false,
    reason: 'disabled.noFormatCopied',
  });
  expect(editor.can({ type: 'setAlign', align: 'ctr' })).toEqual({
    ok: false,
    reason: 'disabled.noText',
  });
  expect(editor.can({ type: 'removeBullet' }).ok).toBe(false);
  editor.destroy();
});

test('typing answers do not carry over to the same text box id on another slide', async () => {
  const editor = await open([textBox(3), textBox(3)]);
  const [box] = ids(editor);
  const at = { paragraph: 0, offset: 0 };
  editor.setSelection({ kind: 'text', nodeId: box!, anchor: at, focus: at });
  editor.input('text', 'a');
  const first = editor.can({ type: 'toggleBold' });
  editor.selectSlide(1);
  const [other] = ids(editor, 1);
  expect(other).toBe(box);
  editor.setSelection({ kind: 'text', nodeId: other!, anchor: at, focus: at });
  expect(editor.can({ type: 'toggleBold' })).not.toBe(first);
  editor.destroy();
});
