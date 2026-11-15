// `ThumbnailPainter.repainted` tells the editor when no thumbnail still waits to repaint, so the
// font faces that the old thumbnails draw with stay registered until then.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import { createPptxEditor } from '../index.ts';
import { ThumbnailPainter } from '../thumbnails.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

test('repainted resolves once the changed thumbnails repaint', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(textDeck(2))).ok).toBe(true);
  const painter = new ThumbnailPainter({
    session: () => editor.session,
    idPrefix: 't',
    frameLabel: (kind) => kind,
    layoutPreview: () => undefined,
    currentSlide: () => 0,
  });
  const box = doc.createElement('div') as unknown as HTMLElement;
  painter.render(1, box, true);
  // Nothing waits: it resolves at once.
  await painter.repainted();

  const lay = editor.session!.layout(1)!;
  const item = lay.items.find((i) => i.kind === 'shape' && i.text)!;
  const caret = { paragraph: 0, offset: 0 };
  editor.goToSlide(1);
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: caret, focus: caret });
  editor.input('text', 'x');
  painter.render(1, box, true);
  let done = false;
  const repainted = painter.repainted().then(() => (done = true));
  await new Promise((r) => setTimeout(r, 0));
  expect(done).toBe(false);
  await repainted;
  expect(box.textContent).toContain('x');
  editor.destroy();
});

test('repainted resolves when the painter is cleared', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(textDeck(2))).ok).toBe(true);
  const painter = new ThumbnailPainter({
    session: () => editor.session,
    idPrefix: 't',
    frameLabel: (kind) => kind,
    layoutPreview: () => undefined,
    currentSlide: () => 0,
  });
  const box = doc.createElement('div') as unknown as HTMLElement;
  painter.render(1, box, true);
  const lay = editor.session!.layout(1)!;
  const item = lay.items.find((i) => i.kind === 'shape' && i.text)!;
  const caret = { paragraph: 0, offset: 0 };
  editor.goToSlide(1);
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: caret, focus: caret });
  editor.input('text', 'y');
  painter.render(1, box, true);
  const repainted = painter.repainted();
  painter.clear();
  await repainted;
  editor.destroy();
});
