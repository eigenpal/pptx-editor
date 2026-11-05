import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { ContextMenu, MenuRow, MenuSeparator, placeInViewport } from '../ContextMenu.tsx';
import { SlideRail } from '../panes.tsx';
import { tooltipText } from './tooltip-text.ts';

afterEach(cleanup);

describe('ContextMenu', () => {
  test('opens at the point, with menu, menuitem, and separator roles', () => {
    render(
      <ContextMenu at={{ x: 40, y: 50 }} onClose={() => undefined} label="Actions">
        <MenuRow label="Copy" icon="content_copy" shortcut="Ctrl+C" onSelect={() => undefined} />
        <MenuSeparator />
        <MenuRow label="Paste" icon="content_paste" onSelect={() => undefined} />
      </ContextMenu>
    );
    const menu = screen.getByRole('menu', { name: 'Actions' });
    expect(menu.style.left).toBe('40px');
    expect(menu.style.top).toBe('50px');
    expect(screen.getAllByRole('menuitem').length).toBe(2);
    expect(screen.getAllByRole('separator').length).toBe(1);
    expect(screen.getByText('Ctrl+C').className).toBe('pptx-menu__shortcut');
  });

  test('stays inside the window: it flips at the right and bottom edges', () => {
    expect(placeInViewport({ x: 990, y: 700 }, 240, 300, 1000, 760)).toEqual({
      left: 750,
      top: 400,
    });
    expect(placeInViewport({ x: 10, y: 10 }, 240, 900, 1000, 760)).toEqual({ left: 10, top: 8 });
  });

  test('arrow keys move through the rows, a click runs one, Escape closes and returns focus', () => {
    const ran: string[] = [];
    let open = true;
    const back = document.createElement('button');
    document.body.append(back);
    back.focus();
    const { rerender } = render(
      <ContextMenu at={{ x: 0, y: 0 }} onClose={() => (open = false)} label="Actions" keyboard>
        <MenuRow label="One" onSelect={() => ran.push('one')} />
        <MenuRow label="Two" onSelect={() => ran.push('two')} />
      </ContextMenu>
    );
    const [one, two] = screen.getAllByRole('menuitem');
    expect(document.activeElement).toBe(one!);
    fireEvent.keyDown(one!, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(two!);
    fireEvent.keyDown(two!, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(one!);
    fireEvent.click(two!);
    expect(ran).toEqual(['two']);
    fireEvent.keyDown(one!, { key: 'Escape' });
    expect(open).toBe(false);
    rerender(<></>);
    expect(document.activeElement).toBe(back);
    back.remove();
  });

  test('a pointer press outside, a wheel outside, a blur, and a resize each close it', () => {
    const outside = document.createElement('div');
    document.body.append(outside);
    const closes: string[] = [];
    const ways: [string, () => void][] = [
      // A press the slide surface prevents fires `pointerdown` and no `mousedown`.
      [
        'pointerdown',
        () => void outside.dispatchEvent(new Event('pointerdown', { bubbles: true })),
      ],
      ['wheel', () => void outside.dispatchEvent(new Event('wheel', { bubbles: true }))],
      ['blur', () => void window.dispatchEvent(new Event('blur'))],
      ['resize', () => void window.dispatchEvent(new Event('resize'))],
    ];
    for (const [name, run] of ways) {
      render(
        <ContextMenu at={{ x: 0, y: 0 }} onClose={() => closes.push(name)} label="Actions">
          <MenuRow label="One" onSelect={() => undefined} />
        </ContextMenu>
      );
      // A press or a wheel inside the menu keeps it.
      const menu = screen.getByRole('menu');
      menu.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      menu.dispatchEvent(new Event('wheel', { bubbles: true }));
      expect(closes.length).toBe(ways.findIndex(([n]) => n === name));
      run();
      expect(closes.at(-1)).toBe(name);
      cleanup();
    }
    outside.remove();
  });

  test('a disabled row shows its reason and does not run', () => {
    const ran: string[] = [];
    render(
      <ContextMenu at={{ x: 0, y: 0 }} onClose={() => undefined} label="Actions">
        <MenuRow
          label="Group"
          disabled
          reason="Select two or more shapes to group"
          onSelect={() => ran.push('group')}
        />
      </ContextMenu>
    );
    const row = screen.getByRole('menuitem', { name: 'Group' });
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(row)).toBe('Select two or more shapes to group');
    fireEvent.click(row);
    expect(ran).toEqual([]);
  });
});

describe('SlideRail context menu', () => {
  async function setup(): Promise<PptxEditorInstance> {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const r = await editor.load(
      buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] })
    );
    expect(r.ok).toBe(true);
    render(
      <PptxEditorContext.Provider value={editor}>
        <SlideRail />
      </PptxEditorContext.Provider>
    );
    return editor;
  }
  const slide = (n: number): HTMLElement => screen.getByRole('option', { name: `Slide ${n}` });

  test('right-click selects an unselected slide, then Duplicate copies it', async () => {
    const editor = await setup();
    fireEvent.contextMenu(slide(3), { clientX: 20, clientY: 30 });
    expect(editor.snapshot().selectedSlides).toEqual([2]);
    const menu = screen.getByRole('menu', { name: 'Slide actions' });
    expect(menu.style.left).toBe('20px');
    fireEvent.click(screen.getByRole('menuitem', { name: 'Duplicate slide' }));
    expect(editor.snapshot().slides.length).toBe(4);
    expect(screen.queryByRole('menu') === null).toBe(true);
  });

  test('right-click inside a multi-selection keeps it and acts on every selected slide', async () => {
    const editor = await setup();
    act(() => {
      editor.selectSlide(0);
      editor.selectSlide(1, 'toggle');
    });
    fireEvent.contextMenu(slide(2), { clientX: 20, clientY: 30 });
    expect(editor.snapshot().selectedSlides).toEqual([0, 1]);
    fireEvent.click(screen.getByRole('menuitem', { name: 'Duplicate 2 slides' }));
    expect(editor.snapshot().slides.length).toBe(5);
  });

  test('Shift+F10 opens the menu on the focused slide; Move up is disabled at the top', async () => {
    await setup();
    fireEvent.keyDown(slide(1), { key: 'F10', shiftKey: true });
    const up = screen.getByRole('menuitem', { name: 'Move up' });
    expect(up.getAttribute('aria-disabled')).toBe('true');
    const down = screen.getByRole('menuitem', { name: 'Move down' });
    expect(down.getAttribute('aria-disabled')).toBeNull();
    expect(document.activeElement).toBe(screen.getAllByRole('menuitem')[0]!);
  });

  test('Delete is disabled with the reason when every slide is selected', async () => {
    const editor = await setup();
    act(() => {
      editor.selectSlide(0);
      editor.selectSlide(2, 'range');
    });
    fireEvent.contextMenu(slide(1), { clientX: 5, clientY: 5 });
    const del = screen.getByRole('menuitem', { name: 'Delete 3 slides' });
    expect(del.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(del)).toContain('A presentation must keep at least one slide');
  });
});
