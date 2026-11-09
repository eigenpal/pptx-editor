// Edits on large content cost in proportion to the content, not to its square: the toolbar's
// state checks, a format change on many shapes, Replace All in a long run, and AutoCorrect after
// a long word. Each case counts work (index copies, resolutions) or compares two sizes, so the
// check does not depend on the machine.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { workingTreeTestRecorder } from '../../store/store/working-tree.ts';
import { corrections } from '../autocorrect.ts';
import { chromeSlotIds, createPptxEditor, toolbarCommandState } from '../index.ts';

const win = new GlobalWindow();
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const NOFIT = '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>';

async function open(shapes: string) {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildPptx({ slides: [{ shapes }] }))).ok).toBe(true);
  return editor;
}

/** Index entries the toolbar's checks copy for a shape of `n` paragraphs. */
async function toolbarIndexWrites(n: number): Promise<number> {
  const text = Array.from({ length: n }, (_, i) => para(run(`p${i}`))).join('');
  const editor = await open(
    shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: 3 * IN, text, bodyPr: NOFIT })
  );
  const id = editor.session!.layout(0)!.items.find((i) => i.kind === 'shape')!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  const recorder = workingTreeTestRecorder();
  recorder.reset();
  for (const slot of chromeSlotIds()) toolbarCommandState(editor, slot);
  const writes = recorder.indexWrites;
  editor.destroy();
  return writes;
}

test('the toolbar checks copy each part index once per check, not once per paragraph', async () => {
  const small = await toolbarIndexWrites(200);
  const large = await toolbarIndexWrites(800);
  // Linear: four times the paragraphs, about four times the work. Quadratic was sixteen.
  expect(large / small).toBeLessThan(6);
});

test('a font change on many shapes resolves the slide once', async () => {
  const shapes = Array.from({ length: 300 }, (_, i) =>
    shape({
      id: i + 2,
      x: (i % 20) * 0.5 * IN,
      y: Math.floor(i / 20) * 0.4 * IN,
      cx: 0.4 * IN,
      cy: 0.3 * IN,
      text: para(run(`s${i}`)),
    })
  ).join('');
  const editor = await open(shapes);
  editor.selectAll();
  const start = performance.now();
  expect(editor.exec({ type: 'setFontFamily', family: 'Carlito' }).ok).toBe(true);
  // A resolution for each shape took seconds; one resolution takes a small part of one.
  expect(performance.now() - start).toBeLessThan(1500);
  editor.destroy();
});

test('Replace All in one long run writes the run once', async () => {
  const text = para(run('ab '.repeat(20_000)));
  const editor = await open(
    shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: 3 * IN, text, bodyPr: NOFIT })
  );
  const recorder = workingTreeTestRecorder();
  recorder.reset();
  const result = editor.find.replaceAll('b', 'c');
  expect(result).toEqual({ ok: true, count: 20_000 });
  expect(editor.find.matches('ac').length).toBe(20_000);
  // One working tree for the slide part, not one for each of the 20000 matches.
  expect(recorder.opens).toBeLessThan(10);
  editor.destroy();
});

test('AutoCorrect after a long word scans it once', () => {
  // No rule applies: the cost is the scan for the word alone.
  const before = `W${'w'.repeat(200_000)}`;
  const start = performance.now();
  expect(corrections(before, ' ').edits).toEqual([]);
  // A linear scan takes milliseconds. The quadratic one took minutes for this length.
  expect(performance.now() - start).toBeLessThan(1000);
});
