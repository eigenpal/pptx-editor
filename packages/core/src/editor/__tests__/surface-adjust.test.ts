// Adjust handles on the canvas (surface-adjust.ts): the diamonds of one selected preset shape,
// and the drag that finds the adjust values for a pointer place. The command they write is in
// adjustments.test.ts.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { apply, presetHandles, type LayoutItem } from '../../layout/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import {
  adjustAt,
  adjustDragLive,
  adjustHandleElements,
  startAdjustDrag,
} from '../surface-adjust.ts';
import { itemById } from '../surface.ts';

const shape = (id: number, name: string, prst: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>` +
  `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const deck = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape(2, 'Arrow', 'rightArrow') + shape(3, 'Box', 'rect'),
      },
    ],
  });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

function item(e: PptxEditorInstance, name: string): LayoutItem {
  const id = e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
  return itemById(e.session!.layout(0)!, id)!;
}

/** The slide place of handle `i` of a preset item. */
function handleAt(it: LayoutItem, i: number): { x: number; y: number } {
  const el = it.element;
  if (el.kind !== 'shape' || el.geometry.kind !== 'preset') throw new Error('no preset');
  const h = presetHandles(el.geometry.preset, el.xfrm.cx, el.xfrm.cy, el.geometry.adjust)[i]!;
  return apply(it.matrix, h.x, h.y);
}

const win = new GlobalWindow();
afterAll(() => win.happyDOM.close());

describe('adjustHandleElements', () => {
  test('draws one labeled diamond per handle, and none for a shape with no handles', async () => {
    const e = await open();
    const doc = win.document as unknown as Document;
    const nodes = adjustHandleElements(doc, item(e, 'Arrow'), 9525, 'Drag');
    expect(nodes.map((n) => n.getAttribute('data-adjust-handle'))).toEqual(['0', '1']);
    expect(nodes.map((n) => n.textContent)).toEqual(['Drag', 'Drag']);
    // The diamond is centered on the handle place.
    const p = handleAt(item(e, 'Arrow'), 1);
    expect(nodes[1]!.getAttribute('d')).toStartWith(`M${p.x} ${p.y - 5 * 9525}L`);
    expect(adjustHandleElements(doc, item(e, 'Box'), 9525, 'Drag')).toEqual([]);
  });
});

describe('adjust drag', () => {
  test('a move finds the values that put the handle under the pointer', async () => {
    const e = await open();
    const it = item(e, 'Arrow');
    const drag = startAdjustDrag(it, 1, 'g1')!;
    expect(drag).not.toBeNull();
    // The pointer on the handle changes nothing.
    expect(adjustAt(drag, handleAt(it, 1))).toBeNull();
    // Move the head handle 0.5 inch left: the head grows.
    const at = handleAt(it, 1);
    const values = adjustAt(drag, { x: at.x - 457_200, y: at.y })!;
    expect(Object.keys(values)).toEqual(['adj2']);
    expect(e.exec({ type: 'setAdjustments', values }).ok).toBe(true);
    const now = handleAt(item(e, 'Arrow'), 1);
    expect(Math.abs(now.x - (at.x - 457_200))).toBeLessThan(2_000);
    // The same place again writes nothing.
    expect(adjustAt(drag, { x: at.x - 457_200, y: at.y })).toBeNull();
  });

  test('a drag goes on only while its shape is the one selected shape', async () => {
    const e = await open();
    const drag = startAdjustDrag(item(e, 'Arrow'), 1, 'g1')!;
    const live = () => adjustDragLive(drag, e.snapshot().selection, e.session!.layout(0)!);
    expect(live()).toBe(true);
    // Another shape selected, or no selection (Escape): the drag ends.
    item(e, 'Box');
    expect(live()).toBe(false);
    e.setSelection({ kind: 'none' });
    expect(live()).toBe(false);
    // The shape resized under the drag: its handle places changed, so the drag ends.
    const it = item(e, 'Arrow');
    expect(live()).toBe(true);
    const { x, y, cx, cy } = it.element.xfrm;
    const nodeId = it.element.nodeId;
    expect(e.exec({ type: 'setTransform', nodeId, x, y, cx: cx * 2, cy, rot: 0 }).ok).toBe(true);
    expect(live()).toBe(false);
  });

  test('a handle index the shape does not have starts no drag', async () => {
    const e = await open();
    expect(startAdjustDrag(item(e, 'Arrow'), 5, 'g1')).toBeNull();
    expect(startAdjustDrag(item(e, 'Box'), 0, 'g1')).toBeNull();
    expect(startAdjustDrag(undefined, 0, 'g1')).toBeNull();
  });
});
