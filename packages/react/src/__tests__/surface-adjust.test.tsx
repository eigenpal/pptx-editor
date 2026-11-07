import { afterEach, describe, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';

// Adjust handles on the slide surface in a DOM: a drag of a diamond is one undo step, and the
// drag ends on Escape, on an undo, and on a move with no button down.

const IN = 914_400;
const PX = 9_525;

const shape = (id: number, name: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${2 * IN}"/></a:xfrm>` +
  '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:sp>';

async function setup(): Promise<{ editor: PptxEditorInstance; host: HTMLDivElement }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  const r = await editor.load(
    buildPptx({ slides: [{ layout: 'blank', shapes: shape(2, 'Round') + shape(3, 'Other') }] })
  );
  expect(r.ok).toBe(true);
  const host = document.createElement('div');
  document.body.append(host);
  editor.attach(host);
  // happy-dom lays nothing out: one CSS pixel is one screen pixel at zoom 1.
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 1280,
      bottom: 720,
      width: 1280,
      height: 720,
      x: 0,
      y: 0,
    }) as DOMRect;
  select(editor, 0);
  return { editor, host };
}

function select(editor: PptxEditorInstance, i: number): void {
  const id = editor.session!.layout(0)!.items[i]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
}

/** The `adj` value of shape `i`, or undefined while it has none of its own. */
const adj = (editor: PptxEditorInstance, i: number): number | undefined => {
  const el = editor.session!.layout(0)!.items[i]!.element;
  return el.kind === 'shape' && el.geometry.kind === 'preset'
    ? el.geometry.adjust.get('adj')
    : undefined;
};

const at = (x: number, y: number) => ({ clientX: x / PX, clientY: y / PX, bubbles: true });
const held = (x: number, y: number) => ({ ...at(x, y), buttons: 1 });

/** Press the diamond of a roundRect at (IN, IN), 4 × 2 in: it sits on the top edge. */
function press(host: HTMLElement): void {
  const diamond = host.querySelector('[data-adjust-handle="0"]')!;
  expect(diamond).not.toBeNull();
  diamond.dispatchEvent(new PointerEvent('pointerdown', { ...at(IN + IN / 3, IN), button: 0 }));
}

const move = (x: number) => document.dispatchEvent(new PointerEvent('pointermove', held(x, IN)));

afterEach(() => {
  document.body.replaceChildren();
});

describe('adjust handles on the surface', () => {
  test('a drag of the diamond is one undo step', async () => {
    const { editor, host } = await setup();
    press(host);
    move(IN + IN / 2);
    move(IN + 0.75 * IN);
    document.dispatchEvent(new PointerEvent('pointerup', at(IN + 0.75 * IN, IN)));
    // The corner radius is 0.75 in of the 2 in short side: adj 37500.
    expect(Math.abs(adj(editor, 0)! - 37_500)).toBeLessThan(600);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(adj(editor, 0)).toBeUndefined();
    expect(editor.exec({ type: 'undo' }).ok).toBe(false);
  });

  test('Escape, an undo, or a move with no button ends the drag', async () => {
    const { editor, host } = await setup();
    // Escape clears the selection: later moves edit nothing and say nothing.
    press(host);
    move(IN + IN / 2);
    const after = adj(editor, 0);
    editor.setSelection({ kind: 'none' });
    move(IN + IN);
    expect(adj(editor, 0)).toBe(after);
    expect(editor.snapshot().lastRejection ?? null).toBeNull();
    document.dispatchEvent(new PointerEvent('pointerup', at(IN + IN, IN)));
    // A selection of another shape during the drag: the other shape keeps its value.
    select(editor, 0);
    press(host);
    select(editor, 1);
    move(IN + IN);
    expect(adj(editor, 1)).toBeUndefined();
    document.dispatchEvent(new PointerEvent('pointerup', at(IN + IN, IN)));
    // An undo during the drag: the next move does not write the value again.
    select(editor, 0);
    press(host);
    move(IN + 0.9 * IN);
    const dragged = adj(editor, 0);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const undone = adj(editor, 0);
    expect(undone).not.toBe(dragged);
    move(IN + 0.95 * IN);
    expect(adj(editor, 0)).toBe(undone);
    document.dispatchEvent(new PointerEvent('pointerup', at(IN + IN, IN)));
    // A lost pointerup: a hover with no button down edits nothing.
    press(host);
    document.dispatchEvent(new PointerEvent('pointermove', at(IN + 0.2 * IN, IN)));
    move(IN + 0.3 * IN);
    expect(adj(editor, 0)).toBe(undone);
  });
});
