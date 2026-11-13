// The drag preview of a shape in a soft-edged group (`p:grpSpPr/a:effectLst/a:softEdge`): the
// group's soft edge filters a copy of its children, and each text paints once in its shape's
// node, so hiding the dragged shape hides its body in the copy and its text, with no ghost.

import { afterAll, describe, expect, spyOn, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  group,
  para,
  run,
  shape,
  solid,
} from '../../../../../scripts/fixtures/builder.ts';
import { resolveSlide } from '../../cascade/index.ts';
import {
  createFixedMeasurer,
  installDrawingRenderers,
  layoutSlide,
  type LayoutItem,
} from '../../layout/index.ts';
import { paintSlide, type PaintOptions } from '../../output/index.ts';
import { readPptx } from '../../store/index.ts';
import { PaintedPreview } from '../drag-preview.ts';
import type { Drag } from '../surface-drag.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;
const OPTS: PaintOptions = { images: () => undefined, idPrefix: 'p' };

describe('drag preview in a soft-edged group', () => {
  test('hiding the dragged shape hides its text; the preview paints it once', () => {
    // A soft edge filter of its own: the paint of the group, not the filter's look, is checked.
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    installDrawingRenderers({
      effects: {
        body: () => ({ filter: 'soft', defs: [] }),
        glow: () => null,
        reflection: () => null,
      },
    });
    warn.mockRestore();
    const deck = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: group(
            {
              id: 10,
              x: IN,
              y: IN,
              cx: 3 * IN,
              cy: IN,
              effects: '<a:effectLst><a:softEdge rad="63500"/></a:effectLst>',
            },
            shape({
              id: 11,
              x: IN,
              y: IN,
              cx: IN,
              cy: IN,
              fill: solid('4472C4'),
              text: para(run('Moved')),
            }) +
              shape({
                id: 12,
                x: 3 * IN,
                y: IN,
                cx: IN,
                cy: IN,
                fill: solid('ED7D31'),
                text: para(run('Kept')),
              })
          ),
        },
      ],
    });
    const r = readPptx(deck);
    if (!r.ok) throw new Error('open failed');
    const layout = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer());
    const slide = paintSlide(doc, layout, OPTS);
    const groupItem = layout.items[0] as Extract<LayoutItem, { kind: 'group' }>;
    const moved = groupItem.children[0]!;
    const id = moved.element.nodeId;
    const shown = (root: Element) =>
      [...root.querySelectorAll('[data-pptx-text]')]
        .filter((t) => !t.closest('[display="none"]'))
        .map((t) => t.textContent);
    expect(shown(slide)).toEqual(['Moved', 'Kept']);
    const preview = new PaintedPreview();
    preview.show(slide, layout, { shown: new Map([[id, moved]]) } as unknown as Drag, OPTS);
    // The slide's own copy of the shape is hidden, text included; the preview paints it once.
    const original = slide.querySelector(`[data-node-id="${id}"]`)!;
    expect(original.getAttribute('display')).toBe('none');
    expect(original.querySelector('[data-pptx-text]')?.textContent).toBe('Moved');
    const previews = [...slide.querySelectorAll('[data-kind="drag-preview"]')];
    expect(previews.length).toBe(1);
    expect(shown(slide).filter((t) => t === 'Moved').length).toBe(1);
    expect(shown(previews[0]!)).toEqual(['Moved']);
    preview.clear();
    expect(original.hasAttribute('display')).toBe(false);
  });
});
