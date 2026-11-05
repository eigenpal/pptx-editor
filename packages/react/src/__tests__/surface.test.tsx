import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { attr, find, xmlPartOf, type Element } from '@pptx-editor.dev/core/store';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { EditNotice } from '../EditNotice.tsx';

// The slide surface in a DOM: its accessible names (E19), the rotation handle of a flipped
// shape (E11), and the edit notice for commands that had nothing to do.

const IN = 914_400;
const PX = 9_525;
const t = createTranslator();

async function setup(): Promise<{ editor: PptxEditorInstance; host: HTMLDivElement }> {
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    zoom: 1,
    translate: (key, params) => t(key as never, params),
  });
  const r = await editor.load(
    buildPptx({
      slides: [
        { shapes: shape({ id: 2, name: 'F', x: IN, y: IN, cx: 2 * IN, cy: IN, flipV: true }) },
        { shapes: '' },
      ],
    })
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
  return { editor, host };
}

const shapeEl = (e: PptxEditorInstance): Element => {
  const part = e.session!.info.slides[0].partName;
  const tree = find(xmlPartOf(e.session!.store.package, part)!.root, 'p:spTree')!;
  return tree.children.find((c): c is Element => c.type === 'element' && c.local === 'sp')!;
};

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

describe('--doc-slide-padding', () => {
  test('sets the room around the fitted slide, 24px a side by default', async () => {
    const css = await Bun.file(
      new URL('../../../core/src/styles/editor.css', import.meta.url)
    ).text();
    const sheet = document.createElement('style');
    sheet.textContent = css;
    document.head.append(sheet);
    const fitFor = async (padding: string | null): Promise<number> => {
      const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
      expect((await editor.load(buildPptx({ slides: [{ shapes: '' }] }))).ok).toBe(true);
      const host = document.createElement('div');
      if (padding) host.style.setProperty('--doc-slide-padding', padding);
      // A 1280 x 720 slide in a host that is 48px wider than the slide.
      Object.defineProperty(host, 'clientWidth', { value: 1328 });
      Object.defineProperty(host, 'clientHeight', { value: 2000 });
      document.body.append(host);
      editor.attach(host);
      const zoom = editor.snapshot().zoom;
      editor.destroy();
      return zoom;
    };
    try {
      expect(await fitFor(null)).toBeCloseTo(1, 3);
      expect(await fitFor('4px')).toBeCloseTo(1320 / 1280, 3);
    } finally {
      sheet.remove();
    }
  });
});

describe('slide surface', () => {
  test('E19: the surface and its text input have accessible names from en.json', async () => {
    const { editor, host } = await setup();
    const root = host.querySelector('.pptx-surface') as HTMLElement;
    expect(root.getAttribute('aria-label')).toBe('Slide 1 of 2');
    expect(root.getAttribute('role')).toBe('group');
    const input = host.querySelector('textarea') as HTMLTextAreaElement;
    expect(input.getAttribute('aria-label')).toBe('Text of the selected shape');
    // The slide is the Tab stop; the hidden text input is not.
    expect(input.tabIndex).toBe(-1);
    editor.goToSlide(1);
    expect(root.getAttribute('aria-label')).toBe('Slide 2 of 2');
    editor.destroy();
  });

  test('every selection mark has a wider halo under it, so it shows on a dark slide', async () => {
    const { editor, host } = await setup();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const overlay = () => host.querySelector('.pptx-surface__overlay')!;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    const outline = overlay().querySelector('.pptx-surface__outline')!;
    const halo = outline.previousElementSibling!;
    expect(halo.getAttribute('class')).toBe('pptx-surface__halo pptx-surface__halo--outline');
    expect(Number(halo.getAttribute('stroke-width'))).toBe(
      3 * Number(outline.getAttribute('stroke-width'))
    );
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    // The text frame is dashed; its halo is solid, so the dashes show dark and light.
    const frame = overlay().querySelector('.pptx-surface__outline')!;
    expect(frame.hasAttribute('stroke-dasharray')).toBe(true);
    expect(frame.previousElementSibling!.hasAttribute('stroke-dasharray')).toBe(false);
    editor.destroy();
  });

  test('a caret in text takes the text color and no halo; an empty line keeps the halo', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    editor.attach(host);
    const white = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>';
    const bytes = buildPptx({
      slides: [
        {
          // A white run, then an empty paragraph with no run to take a color from.
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            text: para(run('Hi', '', white)) + para(''),
          }),
        },
      ],
    });
    await act(async () => void (await editor.load(bytes)));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const caret = (paragraph: number, offset: number) => {
      const at = { paragraph, offset };
      editor.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
      return host.querySelector<SVGLineElement>('.pptx-surface__caret')!;
    };
    const inText = caret(0, 1);
    expect(inText.classList.contains('pptx-surface__caret--text')).toBe(true);
    expect(inText.style.stroke.replace(/\s/g, '')).toMatch(/rgb\(255,?255,?255\)/);
    expect(inText.previousElementSibling?.classList.contains('pptx-surface__halo--caret')).toBe(
      false
    );
    const inEmpty = caret(1, 0);
    expect(inEmpty.classList.contains('pptx-surface__caret--text')).toBe(false);
    expect(inEmpty.previousElementSibling?.classList.contains('pptx-surface__halo--caret')).toBe(
      true
    );
    editor.destroy();
    host.remove();
  });

  test('E11: a drag from the rotation handle of a flipV shape turns it by the pointer move', async () => {
    const { editor, host } = await setup();
    const id = shapeEl(editor).id;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    const handle = host.querySelector('[data-handle="rot"]') as SVGCircleElement;
    const hx = Number(handle.getAttribute('cx'));
    const hy = Number(handle.getAttribute('cy'));
    // The handle is outside the box (y from 1 in to 2 in), not inside it.
    expect(hy < IN || hy > 2 * IN).toBe(true);
    const at = (x: number, y: number) => ({ clientX: x / PX, clientY: y / PX, bubbles: true });
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...at(hx, hy), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(hx + 4 * PX, hy)));
    document.dispatchEvent(new PointerEvent('pointerup', at(hx + 4 * PX, hy)));
    const xfrm = find(find(shapeEl(editor), 'p:spPr'), 'a:xfrm');
    const rot = Number(attr(xfrm, 'rot') ?? 0) / 60_000;
    expect(Math.min(rot, 360 - rot)).toBeGreaterThan(0);
    expect(Math.min(rot, 360 - rot)).toBeLessThan(5);
    expect(attr(xfrm, 'flipV')).toBe('1');
    editor.destroy();
  });

  test('a corner drag keeps the ratio of a shape whose aspect ratio is locked', async () => {
    const { editor, host } = await setup();
    const id = shapeEl(editor).id;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    expect(editor.exec({ type: 'setLockAspectRatio', lock: true }).ok).toBe(true);
    const handle = host.querySelector('[data-handle="se"]') as SVGCircleElement;
    const hx = Number(handle.getAttribute('cx'));
    const hy = Number(handle.getAttribute('cy'));
    const at = (x: number, y: number) => ({ clientX: x / PX, clientY: y / PX, bubbles: true });
    // Drag the corner right only; the height follows (2:1 box).
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...at(hx, hy), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(hx + IN, hy)));
    document.dispatchEvent(new PointerEvent('pointerup', at(hx + IN, hy)));
    const ext = find(find(find(shapeEl(editor), 'p:spPr'), 'a:xfrm'), 'a:ext');
    const cx = Number(attr(ext, 'cx'));
    const cy = Number(attr(ext, 'cy'));
    expect(cx).toBeGreaterThan(2 * IN);
    expect(Math.abs(cx / cy - 2)).toBeLessThan(0.01);
    editor.destroy();
  });

  test('a side drag of a locked shape scales the other axis about its center', async () => {
    const { editor, host } = await setup();
    const id = shapeEl(editor).id;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    const off = () => find(find(find(shapeEl(editor), 'p:spPr'), 'a:xfrm'), 'a:off');
    const ext = () => find(find(find(shapeEl(editor), 'p:spPr'), 'a:xfrm'), 'a:ext');
    const center0 = Number(attr(off(), 'y')) + Number(attr(ext(), 'cy')) / 2;
    expect(editor.exec({ type: 'setLockAspectRatio', lock: true }).ok).toBe(true);
    const handle = host.querySelector('[data-handle="e"]') as SVGCircleElement;
    const hx = Number(handle.getAttribute('cx'));
    const hy = Number(handle.getAttribute('cy'));
    const at = (x: number, y: number) => ({ clientX: x / PX, clientY: y / PX, bubbles: true });
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...at(hx, hy), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(hx + IN, hy)));
    document.dispatchEvent(new PointerEvent('pointerup', at(hx + IN, hy)));
    const cx = Number(attr(ext(), 'cx'));
    const cy = Number(attr(ext(), 'cy'));
    expect(cx).toBeGreaterThan(2 * IN);
    expect(Math.abs(cx / cy - 2)).toBeLessThan(0.01);
    expect(Math.abs(Number(attr(off(), 'y')) + cy / 2 - center0)).toBeLessThan(2);
    editor.destroy();
  });
});

describe('slide surface: drag history', () => {
  const at = (x: number, y: number) => ({ clientX: x / PX, clientY: y / PX, bubbles: true });
  const cx = (e: PptxEditorInstance) =>
    Number(attr(find(find(find(shapeEl(e), 'p:spPr'), 'a:xfrm'), 'a:ext'), 'cx'));

  test('the steps of one pointer drag are one undo step that restores the shape', async () => {
    const { editor, host } = await setup();
    editor.setSelection({ kind: 'shapes', nodeIds: [shapeEl(editor).id] });
    const handle = host.querySelector('[data-handle="e"]') as SVGCircleElement;
    const hx = Number(handle.getAttribute('cx'));
    const hy = Number(handle.getAttribute('cy'));
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...at(hx, hy), button: 0 }));
    for (const step of [1, 2, 3])
      document.dispatchEvent(new PointerEvent('pointermove', at(hx + (step * IN) / 4, hy)));
    document.dispatchEvent(new PointerEvent('pointerup', at(hx + (3 * IN) / 4, hy)));
    expect(cx(editor)).toBeGreaterThan(2 * IN);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(cx(editor)).toBe(2 * IN);
    expect(editor.snapshot().canUndo).toBe(false);
    editor.destroy();
  });

  test('pointercancel ends the drag: a later move changes nothing, and undo is one step', async () => {
    const { editor, host } = await setup();
    editor.setSelection({ kind: 'shapes', nodeIds: [shapeEl(editor).id] });
    const handle = host.querySelector('[data-handle="e"]') as SVGCircleElement;
    const hx = Number(handle.getAttribute('cx'));
    const hy = Number(handle.getAttribute('cy'));
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...at(hx, hy), button: 0 }));
    document.dispatchEvent(new PointerEvent('pointermove', at(hx + IN / 2, hy)));
    document.dispatchEvent(new PointerEvent('pointercancel', at(hx + IN / 2, hy)));
    const width = cx(editor);
    expect(width).toBeGreaterThan(2 * IN);
    document.dispatchEvent(new PointerEvent('pointermove', at(hx + IN, hy)));
    expect(cx(editor)).toBe(width);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(cx(editor)).toBe(2 * IN);
    expect(editor.snapshot().canUndo).toBe(false);
    editor.destroy();
  });
});

describe('EditNotice', () => {
  test('shows no notice when a command had nothing to do, and shows real refusals', async () => {
    const { editor } = await setup();
    render(
      <PptxEditorContext.Provider value={editor}>
        <EditNotice />
      </PptxEditorContext.Provider>
    );
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().lastRejection?.noop).toBe(true);
    expect(screen.queryByRole('status') === null).toBe(true);
    act(() => void editor.exec({ type: 'redo' }));
    expect(screen.queryByRole('status') === null).toBe(true);
    act(() => void editor.exec({ type: 'insertShape', preset: 'noSuchPreset' }));
    expect(screen.getByRole('status').textContent).toContain('noSuchPreset');
    editor.destroy();
  });
});
