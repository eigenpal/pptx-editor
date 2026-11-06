// A diagram drawn from its saved drawing is one read-only object: a press on any of its shapes
// or its text hits the diagram frame, never a shape of the drawing.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { apply, createFixedMeasurer, layoutSlide, type LayoutItem } from '../../layout/index.ts';
import { paintSlide } from '../../output/index.ts';
import { hitItem } from '../surface-hit.ts';
import { Window } from 'happy-dom';

const bytes = new Uint8Array(
  readFileSync(join(import.meta.dirname, '../../../../../fixtures/diagrams.pptx'))
);

function leaves(items: readonly LayoutItem[]): LayoutItem[] {
  return items.flatMap((i) => (i.kind === 'group' ? leaves(i.children) : [i]));
}

describe('a diagram saved drawing in the editor', () => {
  test('a press on each shape of the drawing hits the diagram frame', () => {
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('open failed');
    for (const index of [0, 1]) {
      const layout = layoutSlide(resolveSlide(r.pkg, index)!, createFixedMeasurer());
      const frame = layout.items.find((i) => i.kind === 'frame');
      if (frame?.kind !== 'frame' || !frame.drawing) throw new Error('no drawing');
      const shapes = leaves(frame.drawing);
      expect(shapes.length).toBeGreaterThan(5);
      for (const shape of shapes) {
        const { cx, cy } = shape.element.xfrm;
        const center = apply(shape.matrix, cx / 2, cy / 2);
        expect(hitItem(layout, center, 1)).toBe(frame.element.nodeId);
        // Inside the frame as a scope (a double-click into it), nothing of the drawing is hit.
        expect(hitItem(layout, center, 1, frame.element.nodeId)).not.toBe(shape.element.nodeId);
      }
    }
  });

  test('the painted diagram carries one hit id, on its frame', () => {
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const { document } = new Window();
    const svg = paintSlide(document as unknown as Document, layout, {
      images: () => undefined,
      idPrefix: 't',
    }) as unknown as Element;
    const frames = svg.querySelectorAll('[data-kind="frame"][data-node-id]');
    expect(frames).toHaveLength(1);
    expect(frames[0]!.querySelectorAll('path').length).toBeGreaterThan(5);
    expect(frames[0]!.querySelectorAll('[data-node-id]')).toHaveLength(0);
  });
});
