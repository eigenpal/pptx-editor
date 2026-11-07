// A slide with more items than one frame paints goes on screen over several frames: the first
// frame shows the bottom items, and the rest follow until the slide equals a full paint. The
// editor stays usable in between.

import { afterAll, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../index.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const SHAPES = 4_500;

const deck = () =>
  buildPptx({
    slides: [
      {
        shapes: Array.from({ length: SHAPES }, (_, i) =>
          shape({
            id: i + 2,
            x: (i % 90) * IN * 0.1,
            y: Math.floor(i / 90) * IN * 0.1,
            cx: IN / 20,
            cy: IN / 20,
          })
        ).join(''),
      },
    ],
  });

const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

test('a slide of thousands of shapes paints over several frames, and then in full', async () => {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await editor.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  editor.attach(host);
  const painted = () => host.querySelectorAll('.pptx-surface__slide [data-node-id]').length;
  const ids = editor.session!.layout(0)!.items.map((i) => i.element.nodeId);
  expect(ids.length).toBe(SHAPES);
  // The first frame paints part of the slide, from the bottom of the paint order.
  const first = painted();
  expect(first).toBeGreaterThan(0);
  expect(first).toBeLessThan(SHAPES);
  // A selection in between keeps the painted shapes and the frames that follow.
  editor.setSelection({ kind: 'shapes', nodeIds: [ids[0]!] });
  for (let i = 0; i < 50 && painted() < SHAPES; i += 1) await wait(20);
  expect(painted()).toBe(SHAPES);
  const order = [...host.querySelectorAll('.pptx-surface__slide [data-node-id]')].map((n) =>
    n.getAttribute('data-node-id')
  );
  expect(order).toEqual(ids);
  editor.destroy();
});
