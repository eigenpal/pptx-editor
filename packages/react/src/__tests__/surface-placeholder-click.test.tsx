import { afterEach, describe, expect, test } from 'bun:test';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import type { LayoutItem } from '@pptx-editor.dev/core/layout';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';

// A click in an empty text placeholder or an empty text box places the caret in it, as a click
// in text does, and the caret shows. An empty shape that is neither keeps the move gesture.

const PX = 9_525;
const IN = 914_400;

async function setup(bytes: Uint8Array): Promise<{ editor: PptxEditorInstance; stage: Element }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  expect((await editor.load(bytes)).ok).toBe(true);
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
  return { editor, stage };
}

function clickCenter(stage: Element, item: LayoutItem): void {
  const { x, y, cx, cy } = item.element.xfrm;
  const at = { clientX: (x + cx / 2) / PX, clientY: (y + cy / 2) / PX, bubbles: true };
  stage.dispatchEvent(new PointerEvent('pointerdown', { ...at, button: 0 }));
  document.dispatchEvent(new PointerEvent('pointerup', at));
}

afterEach(() => {
  document.body.replaceChildren();
});

describe('a click in an empty text body', () => {
  test('places the caret in the placeholder, and typed text goes into it', async () => {
    const { editor, stage } = await setup(buildSlideLayoutsDeck());
    editor.goToSlide(4);
    const title = editor.session!.layout(4)!.items.find((i) => i.element.name === 'Title 1')!;
    expect(title.element.placeholder).toBeDefined();
    clickCenter(stage, title);
    const sel = editor.snapshot().selection;
    expect(sel.kind).toBe('text');
    expect(sel.kind === 'text' && sel.nodeId).toBe(title.element.nodeId);
    expect(stage.querySelector('.pptx-surface__caret')).not.toBeNull();
    expect(editor.input('text', 'Q3 plan').ok).toBe(true);
    const after = editor.session!.layout(4)!.items.find((i) => i.element.name === 'Title 1')!;
    const lines = after.kind === 'shape' ? (after.text?.lines ?? []) : [];
    const text = lines
      .flatMap((l) => l.runs.filter((r) => r.kind === 'text').map((r) => r.text))
      .join('');
    expect(text).toBe('Q3 plan');
    editor.destroy();
  });

  test('a click in an empty text box places the caret in it', async () => {
    const { editor, stage } = await setup(buildPptx({ slides: [{ shapes: '' }] }));
    expect(editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
    editor.setSelection({ kind: 'none' });
    const box = editor.session!.layout(0)!.items[0]!;
    expect(box.element.kind === 'shape' && box.element.textBox).toBe(true);
    clickCenter(stage, box);
    expect(editor.snapshot().selection).toMatchObject({ kind: 'text', nodeId: box.element.nodeId });
    expect(stage.querySelector('.pptx-surface__caret')).not.toBeNull();
    editor.destroy();
  });

  test('an empty picture placeholder keeps the move gesture', async () => {
    const ph = (id: number, type: string, x: number): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${id}"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="${IN}"/><a:ext cx="${2 * IN}" cy="${2 * IN}"/></a:xfrm></p:spPr>` +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
    const { editor, stage } = await setup(
      buildPptx({ slides: [{ shapes: ph(2, 'pic', IN) + ph(3, 'tbl', 4 * IN) }] })
    );
    for (const item of editor.session!.layout(0)!.items) {
      editor.setSelection({ kind: 'none' });
      clickCenter(stage, item);
      expect(editor.snapshot().selection).toEqual({
        kind: 'shapes',
        nodeIds: [item.element.nodeId],
      });
    }
    editor.destroy();
  });

  test('in view mode, a click in an empty placeholder places no caret', async () => {
    const { editor, stage } = await setup(buildSlideLayoutsDeck());
    editor.goToSlide(4);
    editor.setMode('view');
    const title = editor.session!.layout(4)!.items.find((i) => i.element.name === 'Title 1')!;
    clickCenter(stage, title);
    expect(editor.snapshot().selection.kind).not.toBe('text');
    expect(stage.querySelector('.pptx-surface__caret')).toBeNull();
    editor.destroy();
  });

  test('an empty shape that is not a placeholder is selected for a move', async () => {
    const bytes = buildPptx({
      slides: [{ shapes: shape({ id: 2, name: 'Box', x: IN, y: IN, cx: 3 * IN, cy: 2 * IN }) }],
    });
    const { editor, stage } = await setup(bytes);
    const box = editor.session!.layout(0)!.items[0]!;
    clickCenter(stage, box);
    expect(editor.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [box.element.nodeId] });
    editor.destroy();
  });
});
