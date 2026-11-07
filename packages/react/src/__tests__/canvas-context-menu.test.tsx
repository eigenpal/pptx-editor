import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditorContent } from '../PptxEditorRoot.tsx';
import { tooltipText } from './tooltip-text.ts';

// The canvas context menu: the surface stops the browser's menu and selects first; the menu
// matches what is under the pointer.

const IN = 914_400;
const PX = 9_525;

async function setup(): Promise<{ editor: PptxEditorInstance; stage: HTMLElement }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  const r = await editor.load(
    buildPptx({
      slides: [
        {
          shapes:
            shape({ id: 2, name: 'A', x: IN, y: IN, cx: 2 * IN, cy: IN }) +
            shape({ id: 3, name: 'B', x: 5 * IN, y: IN, cx: 2 * IN, cy: IN }),
        },
      ],
    })
  );
  expect(r.ok).toBe(true);
  render(
    <PptxEditorContext.Provider value={editor}>
      <PptxEditorContent />
    </PptxEditorContext.Provider>
  );
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  // happy-dom lays nothing out: one CSS pixel is one screen pixel at zoom 1.
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

const at = (x: number, y: number) => ({ clientX: (x * IN) / PX, clientY: (y * IN) / PX });

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

describe('canvas context menu', () => {
  test('right-click on a shape selects it, stops the browser menu, and offers Bring to front', async () => {
    const { editor, stage } = await setup();
    let notPrevented = true;
    act(() => {
      notPrevented = fireEvent.contextMenu(stage, at(2, 1.5));
    });
    expect(notPrevented).toBe(false);
    const sel = editor.snapshot().selection;
    expect(sel.kind === 'shapes' && sel.nodeIds.length).toBe(1);
    screen.getByRole('menu', { name: 'Shape actions' });
    const front = screen.getByRole('menuitem', { name: 'Bring to front' });
    expect(front.getAttribute('aria-disabled')).toBeNull();
    // Without the pro comments pane, the menu has no New comment row.
    expect(document.querySelector('[role="menu"] [data-slot="review.addComment"]')).toBeNull();
    fireEvent.click(front);
    expect(screen.queryByRole('menu') === null).toBe(true);
    expect(editor.snapshot().canUndo).toBe(true);
  });

  test('right-click inside a multi-selection keeps it', async () => {
    const { editor, stage } = await setup();
    act(() => editor.selectAll());
    act(() => {
      fireEvent.contextMenu(stage, at(6, 1.5));
    });
    const sel = editor.snapshot().selection;
    expect(sel.kind === 'shapes' && sel.nodeIds.length).toBe(2);
    const group = screen.getByRole('menuitem', { name: 'Group' });
    expect(group.getAttribute('aria-disabled')).toBeNull();
  });

  test('right-click on empty slide area clears the selection and offers New slide', async () => {
    const { editor, stage } = await setup();
    act(() => editor.selectAll());
    act(() => {
      fireEvent.contextMenu(stage, at(10, 6));
    });
    expect(editor.snapshot().selection.kind).toBe('none');
    screen.getByRole('menu', { name: 'Slide actions' });
    const paste = screen.getByRole('menuitem', { name: 'Paste' });
    expect(paste.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(paste)).toContain('Copy a shape first');
    fireEvent.click(screen.getByRole('menuitem', { name: 'New slide' }));
    expect(editor.snapshot().slides.length).toBe(2);
  });

  test('Copy then Paste from the menu uses the editor clipboard', async () => {
    const { editor, stage } = await setup();
    act(() => {
      fireEvent.contextMenu(stage, at(2, 1.5));
    });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Copy' }));
    expect(editor.lastCopied !== null).toBe(true);
    act(() => {
      fireEvent.contextMenu(stage, at(10, 6));
    });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Paste' }));
    expect(editor.snapshot().selection.kind).toBe('shapes');
  });
});
