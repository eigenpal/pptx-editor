// The text selection highlight (surface-text-selection.ts): one rectangle for each selected
// line, each with a finite position and a visible width.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type LayoutItem } from '../../layout/index.ts';
import { selectionRects } from '../surface-text-selection.ts';

const IN = 914_400;
const WORDS = 'selected text that wraps onto more than one line in the box '.repeat(3);

function item(pPr: string): LayoutItem {
  const r = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: 2 * IN,
            text: para(run(WORDS), pPr),
          }),
        },
      ],
    })
  );
  if (!r.ok) throw new Error('open failed');
  return layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).items[0];
}

function rects(pPr: string): { x: number; width: number }[] {
  const doc = new Window().document as unknown as Document;
  const it = item(pPr);
  if (it.kind !== 'shape' || !it.text) throw new Error('no text');
  const lines = it.text.lines;
  expect(lines.length).toBeGreaterThan(1);
  const [g] = selectionRects(
    doc,
    it,
    { paragraph: 0, offset: 2 },
    { paragraph: 0, offset: lines[1].start + 3 },
    'sel'
  );
  return [...g.querySelectorAll('rect')].map((r) => ({
    x: Number(r.getAttribute('x')),
    width: Number(r.getAttribute('width')),
  }));
}

describe('text selection highlight', () => {
  for (const [name, pPr] of [
    ['left-aligned', '<a:pPr/>'],
    ['justified', '<a:pPr algn="just"/>'],
    ['distributed', '<a:pPr algn="dist"/>'],
  ] as const)
    test(`a ${name} selection over two lines draws both lines`, () => {
      const boxes = rects(pPr);
      expect(boxes.length).toBe(2);
      for (const box of boxes) {
        expect(Number.isFinite(box.x)).toBe(true);
        expect(Number.isFinite(box.width)).toBe(true);
        expect(box.width).toBeGreaterThan(0);
      }
    });
});
