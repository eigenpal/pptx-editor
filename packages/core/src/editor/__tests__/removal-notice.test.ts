// `snapshot.lastRemoval`: a Delete of shapes or slides reports what went, so chrome can offer
// Undo. The next edit, an undo, and a load clear it. Cut sets none
// (`packages/vue/test/overlays.test.ts`, which cuts through a real surface).

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('A')) }) +
          shape({ id: 3, x: 3 * IN, y: IN, cx: IN, cy: IN }),
      },
      { shapes: '' },
      { shapes: '' },
    ],
  });

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  return e;
}

const shapeIds = (e: PptxEditorInstance): string[] =>
  e.session!.layout(0)!.items.map((i) => i.element.nodeId);

describe('lastRemoval', () => {
  test('a slide delete reports one slide; deleting two reports two', async () => {
    const e = await open();
    expect(e.snapshot().lastRemoval).toBeNull();
    expect(e.exec({ type: 'deleteSlide', index: 2 }).ok).toBe(true);
    const first = e.snapshot().lastRemoval!;
    expect(first).toMatchObject({ kind: 'slides', count: 1 });
    expect(e.exec({ type: 'deleteSlides', indices: [0, 1] }).ok).toBe(false); // the last slide stays
    expect(e.snapshot().lastRemoval).toEqual(first);
    e.exec({ type: 'undo' });
    expect(e.snapshot().lastRemoval).toBeNull();
    expect(e.exec({ type: 'deleteSlides', indices: [1, 2] }).ok).toBe(true);
    const second = e.snapshot().lastRemoval!;
    expect(second).toMatchObject({ kind: 'slides', count: 2 });
    expect(second.serial).toBeGreaterThan(first.serial);
  });

  test('a shape delete reports the selected count; the next edit clears it', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: shapeIds(e) });
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    expect(e.snapshot().lastRemoval).toMatchObject({ kind: 'shapes', count: 2 });
    expect(e.exec({ type: 'insertShape', preset: 'rect' }).ok).toBe(true);
    expect(e.snapshot().lastRemoval).toBeNull();
  });

  test('a slide named twice in a delete counts once', async () => {
    const e = await open();
    expect(e.exec({ type: 'deleteSlides', indices: [1, 1] }).ok).toBe(true);
    expect(e.snapshot().lastRemoval).toMatchObject({ kind: 'slides', count: 1 });
  });

  test('opening a document clears it', async () => {
    const e = await open();
    e.exec({ type: 'deleteSlide', index: 1 });
    expect(e.snapshot().lastRemoval).not.toBeNull();
    expect((await e.load(deck())).ok).toBe(true);
    expect(e.snapshot().lastRemoval).toBeNull();
  });
});
