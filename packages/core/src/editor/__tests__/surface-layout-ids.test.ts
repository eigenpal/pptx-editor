// Node ids are unique in one part only: a layout shape and a slide shape at the same position
// in their parts have the same node id, whatever their `cNvPr` ids are. A press on the slide
// shape selects it, the selection marks its box, and a drag previews it, not the layout shape.
// The canvas around the slide acts as the slide: a drag there selects the shapes it encloses, a
// click or a right-click clears the selection, and the pointer shows the move cursor on a
// selected shape's edge past the slide.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { findItem } from '../../automation/commands/placement.ts';
import { createPptxEditor } from '../index.ts';

const IN = 914_400;
const STAGE = { left: 100, top: 100, width: 1280 };

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

/**
 * A layout header text at the top left, and a slide title below it at the same position in its
 * part, with another `cNvPr` id. A box past the right edge of the slide.
 */
function deck(): Uint8Array {
  return buildPptx({
    layouts: [
      {
        name: 'Branded',
        shapes: shape({
          id: 9,
          name: 'Header',
          x: IN / 4,
          y: IN / 8,
          cx: IN,
          cy: IN / 4,
          text: para(run('Brand')),
        }),
      },
    ],
    slides: [
      {
        layoutName: 'Branded',
        shapes:
          shape({
            id: 2,
            name: 'Title',
            x: IN / 2,
            y: IN,
            cx: 6 * IN,
            cy: IN / 2,
            text: para(run('Heading')),
          }) +
          shape({ id: 3, name: 'Box', x: 4 * IN, y: 3 * IN, cx: IN, cy: IN, fill: '336699' }) +
          shape({ id: 4, name: 'Edge', x: 13 * IN, y: 3 * IN, cx: IN, cy: IN, fill: '336699' }),
      },
    ],
  });
}

async function attached() {
  const g = globalThis as { ResizeObserver?: unknown };
  g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await e.load(deck())).ok).toBe(true);
  const host = doc.createElement('div') as unknown as HTMLElement;
  doc.body.append(host as unknown as Node);
  e.attach(host);
  // happy-dom lays nothing out: without a zoom, the fit zoom is off.
  e.setZoom(1);
  const root = host.querySelector('.pptx-surface') as HTMLElement;
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  const width = e.session!.layout(0)!.slide.width;
  const scale = STAGE.width / width;
  stage.getBoundingClientRect = () =>
    ({
      left: STAGE.left,
      top: STAGE.top,
      width: STAGE.width,
      height: 720,
      right: STAGE.left + STAGE.width,
      bottom: STAGE.top + 720,
    }) as DOMRect;
  const mouse = (type: string, x: number, y: number) =>
    new win.MouseEvent(type, {
      clientX: STAGE.left + x * scale,
      clientY: STAGE.top + y * scale,
      button: 0,
      bubbles: true,
      cancelable: true,
    }) as unknown as Event;
  const named = (name: string) =>
    e
      .session!.layout(0)!
      .items.find((i) => i.element.layer === 'slide' && i.element.name === name)!;
  return { e, host, root, stage, mouse, named };
}

describe('a slide shape with the node id of a layout shape', () => {
  test('the slide shape wins the id lookup', async () => {
    const { e } = await attached();
    const { items } = e.session!.layout(0)!;
    const title = items.find((i) => i.element.layer === 'slide' && i.element.name === 'Title')!;
    const same = items.filter((i) => i.element.nodeId === title.element.nodeId);
    // Both shapes are laid out: the layout's header paints first, behind the slide's shapes.
    expect(same.map((i) => i.element.layer)).toEqual(['layout', 'slide']);
    expect(findItem(items, title.element.nodeId)).toBe(title);
    e.destroy();
  });

  test('a press on the slide shape selects it and outlines its own box', async () => {
    const { e, host, stage, mouse, named } = await attached();
    const title = named('Title');
    stage.dispatchEvent(mouse('pointerdown', 2 * IN, IN + IN / 4));
    doc.dispatchEvent(mouse('pointerup', 2 * IN, IN + IN / 4));
    const sel = e.snapshot().selection;
    expect(sel.kind === 'text' || sel.kind === 'shapes').toBe(true);
    expect(sel.kind === 'text' ? sel.nodeId : sel.kind === 'shapes' ? sel.nodeIds[0] : '').toBe(
      title.element.nodeId
    );
    // The outline sits on the title's box, not on the layout header's box.
    const outline = host.querySelector('.pptx-surface__overlay .pptx-surface__outline')!;
    const g = outline.parentElement!;
    expect(g.getAttribute('transform')).toBe(`matrix(${title.matrix.join(' ')})`);
    e.destroy();
  });

  test('a drag hides and previews the slide shape, not the layout shape', async () => {
    const { e, host, stage, mouse, named } = await attached();
    const id = named('Title').element.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    // A press on the title's top edge moves it.
    stage.dispatchEvent(mouse('pointerdown', 3 * IN, IN));
    doc.dispatchEvent(mouse('pointermove', 3 * IN, 2 * IN));
    const hidden = [...host.querySelectorAll(`[data-node-id="${id}"][display="none"]`)];
    expect(hidden.map((el) => el.getAttribute('data-layer'))).toEqual(['slide']);
    expect(host.querySelector(`[data-preview-of="${id}"]`)).not.toBeNull();
    doc.dispatchEvent(mouse('pointerup', 3 * IN, 2 * IN));
    expect(named('Title').element.xfrm.y).toBe(2 * IN);
    e.destroy();
  });
});

describe('a press on the canvas around the slide', () => {
  test('starts a drag selection that selects the shapes it encloses', async () => {
    const { e, root, mouse, named } = await attached();
    // From above and left of the slide, to just past the box.
    root.dispatchEvent(mouse('pointerdown', -IN / 4, -IN / 4));
    doc.dispatchEvent(mouse('pointermove', 7 * IN, 4.5 * IN));
    doc.dispatchEvent(mouse('pointerup', 7 * IN, 4.5 * IN));
    const sel = e.snapshot().selection;
    expect(sel.kind).toBe('shapes');
    // The title and the box, not the layout header, which a slide does not edit.
    expect(sel.kind === 'shapes' ? [...sel.nodeIds].sort() : []).toEqual(
      [named('Title').element.nodeId, named('Box').element.nodeId].sort()
    );
    e.destroy();
  });

  test('a click there clears the selection', async () => {
    const { e, root, mouse } = await attached();
    const ids = e.snapshot().shapes!.map((s) => s.nodeId);
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    root.dispatchEvent(mouse('pointerdown', -IN / 4, -IN / 4));
    doc.dispatchEvent(mouse('pointerup', -IN / 4, -IN / 4));
    expect(e.snapshot().selection.kind).toBe('none');
    e.destroy();
  });

  test('a right-click there clears the selection before the slide menu opens', async () => {
    const { e, root, mouse, named } = await attached();
    e.setSelection({ kind: 'shapes', nodeIds: [named('Box').element.nodeId] });
    const menu = mouse('contextmenu', -IN / 4, -IN / 4);
    root.dispatchEvent(menu);
    expect(menu.defaultPrevented).toBe(true);
    expect(e.snapshot().selection.kind).toBe('none');
    e.destroy();
  });

  test('the pointer shows the move cursor on a selected shape past the slide', async () => {
    // The hover reads `instanceof Node`; the engine test run registers no global DOM.
    const g = globalThis as { Node?: unknown };
    const had = 'Node' in g;
    g.Node ??= (win as unknown as { Node: unknown }).Node;
    try {
      const { e, root, stage, mouse, named } = await attached();
      e.setSelection({ kind: 'shapes', nodeIds: [named('Edge').element.nodeId] });
      // The right edge of the box, past the slide's right edge.
      root.dispatchEvent(mouse('pointermove', 14 * IN, 3.5 * IN));
      expect(root.style.cursor).toBe('move');
      // Back over the slide, the canvas lets go of the cursor: the slide shows its own.
      stage.dispatchEvent(mouse('pointermove', 2 * IN, 5 * IN));
      expect(root.style.cursor).toBe('');
      // Away from every shape, the canvas sets no cursor: the default pointer shows.
      root.dispatchEvent(mouse('pointermove', -IN / 4, -IN / 4));
      expect(root.style.cursor).toBe('');
      e.destroy();
    } finally {
      if (!had) delete g.Node;
    }
  });
});
