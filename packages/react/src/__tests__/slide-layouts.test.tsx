// The slide chrome of G2 (`editor-ui-parity` 2.1–2.3): the layout gallery, Layout, Reset,
// and Hide slide in the rail's context menu, New slide with a layout, and the `slide.*` slots.

import { afterEach, beforeEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen, within } from '@testing-library/react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { SlideRail, ZoomPill } from '../panes.tsx';
import { LayoutGallery } from '../SlideControls.tsx';
import { Toolbar, ToolbarSlot } from '../Toolbar.tsx';
import { tooltipText } from './tooltip-text.ts';

afterEach(cleanup);

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

const layoutName = (editor: PptxEditorInstance, index: number): string | undefined => {
  const id = editor.snapshot().slides[index]?.layoutId;
  return editor.snapshot().layouts?.find((l) => l.id === id)?.name;
};

async function rail(): Promise<PptxEditorInstance> {
  const editor = await open();
  render(
    <PptxEditorContext.Provider value={editor}>
      <SlideRail />
    </PptxEditorContext.Provider>
  );
  return editor;
}

/** A slide of the rail: an option in a flat rail, a tree item in a rail with sections. */
const slide = (name: string): HTMLElement =>
  screen.queryByRole('option', { name }) ?? screen.getByRole('treeitem', { name });
/** A gallery tile's name (its thumbnail is hidden from the accessibility tree). */
const tileName = (tile: HTMLElement): string | null | undefined =>
  tile.querySelector('.pptx-layout-gallery__name')?.textContent;

describe('hidden slides in the rail', () => {
  test('a hidden slide is named hidden and marked for styling', async () => {
    await rail();
    const hidden = slide('Slide 4 (hidden)');
    expect(hidden.hasAttribute('data-hidden')).toBe(true);
    expect(slide('Slide 3').hasAttribute('data-hidden')).toBe(false);
  });

  test('Hide slide hides the selected slides; on a hidden slide it shows them again', async () => {
    const editor = await rail();
    fireEvent.contextMenu(slide('Slide 2'), { clientX: 10, clientY: 10 });
    const hide = screen.getByRole('menuitemcheckbox', { name: 'Hide slide' });
    expect(hide.getAttribute('aria-checked')).toBe('false');
    fireEvent.click(hide);
    expect(editor.snapshot().slides[1]!.hidden).toBe(true);
    fireEvent.contextMenu(slide('Slide 2 (hidden)'), { clientX: 10, clientY: 10 });
    const show = screen.getByRole('menuitemcheckbox', { name: 'Hide slide' });
    expect(show.getAttribute('aria-checked')).toBe('true');
    fireEvent.click(show);
    expect(editor.snapshot().slides[1]!.hidden).toBe(false);
  });
});

describe('LayoutGallery props', () => {
  test('layouts limits the tiles, and disabled or viewing mode takes no pick', async () => {
    const editor = await open();
    const all = editor.snapshot().layouts!;
    const picked: string[] = [];
    const { rerender } = render(
      <PptxEditorContext.Provider value={editor}>
        <LayoutGallery layouts={all.slice(0, 2)} onChange={(id) => picked.push(id)} />
      </PptxEditorContext.Provider>
    );
    const tiles = () => [...document.querySelectorAll<HTMLElement>('.pptx-layout-gallery__item')];
    expect(tiles().length).toBe(2);
    fireEvent.click(tiles()[0]!);
    expect(picked).toEqual([all[0]!.id]);
    rerender(
      <PptxEditorContext.Provider value={editor}>
        <LayoutGallery disabled onChange={(id) => picked.push(id)} />
      </PptxEditorContext.Provider>
    );
    expect(tiles().length).toBe(all.length);
    expect(tiles()[0]!.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(tiles()[0]!);
    act(() => editor.setMode('view'));
    rerender(
      <PptxEditorContext.Provider value={editor}>
        <LayoutGallery onChange={(id) => picked.push(id)} />
      </PptxEditorContext.Provider>
    );
    fireEvent.click(tiles()[1]!);
    expect(picked.length).toBe(1);
    expect(tooltipText(tiles()[1]!)).toContain('open for viewing');
    // disabledReason takes a catalog key, which shows translated, or plain text.
    rerender(
      <PptxEditorContext.Provider value={editor}>
        <LayoutGallery disabled disabledReason="disabled.noSelection" onChange={() => {}} />
      </PptxEditorContext.Provider>
    );
    expect(tooltipText(tiles()[0]!)).toContain('Select a shape first');
    expect(tooltipText(tiles()[0]!)).not.toContain('disabled.');
    rerender(
      <PptxEditorContext.Provider value={editor}>
        <LayoutGallery disabled disabledReason="Pick a slide first" onChange={() => {}} />
      </PptxEditorContext.Provider>
    );
    expect(tooltipText(tiles()[0]!)).toContain('Pick a slide first');
    // Each tile names the reason through aria-describedby, as other disabled chrome does.
    const said = (tile: HTMLElement) =>
      document.getElementById(tile.getAttribute('aria-describedby') ?? '')?.textContent;
    expect(tiles().map(said)).toEqual(tiles().map(() => 'Pick a slide first'));
  });
});

describe('the zoom pill', () => {
  test('is a named group, and its level button says what it is', async () => {
    const editor = await open();
    render(
      <PptxEditorContext.Provider value={editor}>
        <ZoomPill />
      </PptxEditorContext.Provider>
    );
    expect(screen.getByRole('group', { name: 'Zoom' }) !== null).toBe(true);
    const level = screen.getByRole('button', { name: /^Zoom: \d+%$/ });
    expect(tooltipText(level)).toBe(level.getAttribute('aria-label') ?? '');
  });
});

describe('rail keys', () => {
  test('Ctrl+M adds a slide; Cmd+M and Cmd+Ctrl+M do not', async () => {
    const editor = await rail();
    const count = editor.snapshot().slides.length;
    const first = slide('Slide 1');
    act(() => first.focus());
    fireEvent.keyDown(first, { key: 'm', metaKey: true });
    fireEvent.keyDown(first, { key: 'm', metaKey: true, ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(count);
    fireEvent.keyDown(first, { key: 'm', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(count + 1);
  });
});

describe('the New slide key on the canvas', () => {
  test('Ctrl+M adds a slide while text is edited; Cmd+M does not', async () => {
    const editor = await open();
    const host = document.createElement('div');
    document.body.append(host);
    editor.attach(host);
    const nodeId = editor.session!.layout(0)!.items.find((i) => i.kind === 'shape' && i.text)!
      .element.nodeId;
    const at = { paragraph: 0, offset: 0 };
    editor.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
    const count = editor.snapshot().slides.length;
    // The keys reach the surface's text field, as typing does.
    const input = host.querySelector('textarea')!;
    const key = (init: KeyboardEventInit) =>
      input.dispatchEvent(
        new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init })
      );
    key({ key: 'm', metaKey: true });
    expect(editor.snapshot().slides.length).toBe(count);
    key({ key: 'm', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(count + 1);
    editor.destroy();
  });
});

describe('rail focus', () => {
  test('Delete on the last slide, and an undo that removes one, keep focus in the rail', async () => {
    const editor = await rail();
    const count = editor.snapshot().slides.length;
    act(() => editor.goToSlide(count - 1));
    // A rail with sections is a tree: its slides are tree items, not options.
    const last = slide(`Slide ${count}`);
    act(() => last.focus());
    fireEvent.keyDown(last, { key: 'Delete' });
    expect(editor.snapshot().slides.length).toBe(count - 1);
    const now = document.activeElement as HTMLElement;
    expect(now.getAttribute('aria-current')).toBe('true');
    // The next key still acts on the rail: Ctrl+Z brings the slide back.
    fireEvent.keyDown(now, { key: 'z', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(count);
    // Ctrl+D on the last slide, then Ctrl+Z, which removes the focused copy.
    act(() => editor.goToSlide(count - 1));
    const end = slide(`Slide ${count}`);
    act(() => end.focus());
    fireEvent.keyDown(end, { key: 'd', ctrlKey: true });
    fireEvent.keyDown(document.activeElement!, { key: 'z', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(count);
    expect((document.activeElement as HTMLElement).classList.contains('pptx-rail__item')).toBe(
      true
    );
  });
});

describe('Layout and Reset in the rail context menu', () => {
  test('the Layout submenu marks the current layout and changes the selected slides', async () => {
    const editor = await rail();
    act(() => {
      editor.selectSlide(1);
      editor.selectSlide(2, 'range');
    });
    fireEvent.contextMenu(slide('Slide 2'), { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Layout' }));
    const panel = screen.getByRole('menu', { name: 'Layout' });
    const tiles = within(panel).getAllByRole('menuitemradio');
    expect(tiles.map(tileName)).toEqual([
      'Title and Content',
      'Blank',
      'Title Slide',
      'Two Content',
      'Title Only',
      'Quote',
    ]);
    // The current slide (slide 3, Shift+clicked last) is on Two Content.
    expect(tiles.map((t) => t.getAttribute('aria-checked'))).toEqual([
      'false',
      'false',
      'false',
      'true',
      'false',
      'false',
    ]);
    // Each tile is the engine's render of a new slide on that layout.
    expect(tiles[2]!.querySelector('.pptx-layout-gallery__thumb svg') !== null).toBe(true);
    fireEvent.click(tiles[4]!);
    expect([layoutName(editor, 1), layoutName(editor, 2)]).toEqual(['Title Only', 'Title Only']);
    expect(screen.queryByRole('menu') === null).toBe(true);
    act(() => void editor.exec({ type: 'undo' }));
    expect([layoutName(editor, 1), layoutName(editor, 2)]).toEqual([
      'Title and Content',
      'Two Content',
    ]);
  });

  test('Reset slide resets the selected slide', async () => {
    const editor = await rail();
    fireEvent.contextMenu(slide('Slide 2'), { clientX: 10, clientY: 10 });
    const revision = editor.snapshot().revision;
    fireEvent.click(screen.getByRole('menuitem', { name: 'Reset slide' }));
    expect(editor.snapshot().revision).toBeGreaterThan(revision);
    expect(editor.snapshot().canUndo).toBe(true);
  });

  test('viewing mode disables Layout, Reset, and Hide with the reason', async () => {
    const editor = await rail();
    act(() => editor.setMode('view'));
    fireEvent.contextMenu(slide('Slide 1'), { clientX: 10, clientY: 10 });
    for (const row of [
      screen.getByRole('menuitem', { name: 'Layout' }),
      screen.getByRole('menuitem', { name: 'Reset slide' }),
      screen.getByRole('menuitemcheckbox', { name: 'Hide slide' }),
    ])
      expect(row.getAttribute('aria-disabled')).toBe('true');
  });
});

describe('New slide with a layout', () => {
  test('the rail’s layout button adds a slide on the picked layout', async () => {
    const editor = await rail();
    act(() => editor.goToSlide(2));
    fireEvent.click(screen.getByRole('button', { name: 'Add slide with layout' }));
    fireEvent.click(screen.getByRole('menuitem', { name: 'Quote' }));
    expect(editor.snapshot().slides.length).toBe(7);
    expect(editor.snapshot().currentSlide).toBe(3);
    expect(layoutName(editor, 3)).toBe('Quote');
  });

  test('the rail’s layout button says it opens a menu, and a second click closes it', async () => {
    await rail();
    const button = screen.getByRole('button', { name: 'Add slide with layout' });
    expect(button.getAttribute('aria-haspopup')).toBe('menu');
    expect(button.getAttribute('aria-expanded')).toBe('false');
    fireEvent.mouseDown(button);
    fireEvent.click(button);
    expect(button.getAttribute('aria-expanded')).toBe('true');
    expect(screen.getAllByRole('menuitem').length).toBe(6);
    fireEvent.mouseDown(button);
    fireEvent.click(button);
    expect(button.getAttribute('aria-expanded')).toBe('false');
    expect(screen.queryAllByRole('menuitem').length).toBe(0);
  });

  test('the gallery lists nothing to pick before a document opens', () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    render(
      <PptxEditorContext.Provider value={editor}>
        <LayoutGallery onChange={() => undefined} />
      </PptxEditorContext.Provider>
    );
    expect(screen.getByText('This presentation has no layouts.') !== null).toBe(true);
  });
});

describe('the slide.* slots', () => {
  test('Layout on a slide whose layout the master does not list marks no tile', async () => {
    const f = unzipSync(buildSlideLayoutsDeck());
    const master = 'ppt/slideMasters/slideMaster1.xml';
    f[master] = strToU8(
      strFromU8(f[master]!).replace(/<p:sldLayoutId id="\d+" r:id="rIdLayout6"\/>/, '')
    );
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(zipSync(f))).ok).toBe(true);
    act(() => editor.goToSlide(5));
    render(
      <PptxEditorContext.Provider value={editor}>
        <ToolbarSlot slotId="slide.layout" />
      </PptxEditorContext.Provider>
    );
    const trigger = screen.getByRole('button', { name: 'Layout' });
    expect(trigger.hasAttribute('aria-disabled')).toBe(false);
    fireEvent.click(trigger);
    const tiles = screen.getAllByRole('menuitemradio');
    expect(tiles.length).toBeGreaterThan(0);
    expect(tiles.filter((t) => t.getAttribute('aria-checked') === 'true')).toEqual([]);
  });

  test('Toolbar.Slot renders New slide, Layout, Reset slide, and Hide slide', async () => {
    const editor = await open();
    act(() => editor.goToSlide(3));
    render(
      <PptxEditorContext.Provider value={editor}>
        <ToolbarSlot slotId="slide.add" />
        <ToolbarSlot slotId="slide.layout" />
        <ToolbarSlot slotId="slide.reset" />
        <ToolbarSlot slotId="slide.hide" />
      </PptxEditorContext.Provider>
    );
    // The current slide is hidden: Hide slide shows as pressed, and a press shows it.
    const hide = screen.getByRole('button', { name: 'Hide slide' });
    expect(hide.getAttribute('aria-pressed')).toBe('true');
    fireEvent.click(hide);
    expect(editor.snapshot().slides[3]!.hidden).toBe(false);
    fireEvent.click(screen.getByRole('button', { name: 'Layout' }));
    const current = screen
      .getAllByRole('menuitemradio')
      .find((t) => t.getAttribute('aria-checked') === 'true');
    expect(current && tileName(current)).toBe('Title Only');
    fireEvent.click(screen.getByRole('menuitemradio', { name: 'Two Content' }));
    expect(layoutName(editor, 3)).toBe('Two Content');
    fireEvent.click(screen.getByRole('button', { name: 'Choose a layout' }));
    fireEvent.click(screen.getByRole('menuitem', { name: 'Title Slide' }));
    expect(layoutName(editor, 4)).toBe('Title Slide');
    fireEvent.click(screen.getByRole('button', { name: /^New slide/ }));
    // After a title slide, New Slide takes Title and Content.
    expect(layoutName(editor, 5)).toBe('Title and Content');
  });
});

describe('the slide.* slots on the customization ladder', () => {
  let warn: ReturnType<typeof spyOn<Console, 'warn'>>;
  beforeEach(() => {
    resetDevWarnings();
    warn = spyOn(console, 'warn').mockImplementation(() => {});
  });
  afterEach(() => warn.mockRestore());

  const mount = (editor: PptxEditorInstance, ui: React.ReactNode) =>
    render(<PptxEditorContext.Provider value={editor}>{ui}</PptxEditorContext.Provider>);

  test('step 1: className and state attributes', async () => {
    const editor = await open();
    act(() => editor.goToSlide(3));
    const { container } = mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.add" className="acme-new" />
        <ToolbarSlot slotId="slide.layout" className="acme-layout" />
        <ToolbarSlot slotId="slide.hide" className="acme-hide" />
      </Toolbar>
    );
    const at = (id: string) => container.querySelector<HTMLElement>(`[data-slot="${id}"]`)!;
    expect(at('slide.add').classList.contains('acme-new')).toBe(true);
    expect(at('slide.layout').classList.contains('acme-layout')).toBe(true);
    expect(at('slide.hide').hasAttribute('data-active')).toBe(true);
    // The New slide caret keeps its own class: the host class goes to the main button only.
    expect(container.querySelectorAll('.acme-new').length).toBe(1);
    act(() => editor.setMode('view'));
    expect(at('slide.add').hasAttribute('data-disabled')).toBe(true);
    expect(tooltipText(screen.getByRole('button', { name: 'Choose a layout' }))).toContain(
      'The presentation is open for viewing'
    );
  });

  test('steps 2 and 3: New slide takes an icon and asChild without a warning', async () => {
    const editor = await open();
    mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.add" icon={<span data-testid="acme-icon" />} />
        <ToolbarSlot slotId="slide.reset" asChild>
          <a href="#reset" data-testid="acme-reset" />
        </ToolbarSlot>
      </Toolbar>
    );
    expect(screen.getByTestId('acme-icon').closest('[data-slot="slide.add"]') !== null).toBe(true);
    const reset = screen.getByTestId('acme-reset');
    expect(reset.getAttribute('aria-label')).toBe('Reset slide');
    fireEvent.click(reset);
    expect(editor.snapshot().canUndo).toBe(true);
    expect(warn.mock.calls.length).toBe(0);
  });

  test('steps 2 and 3 on Layout: an icon, and asChild on a host element, open the gallery', async () => {
    const editor = await open();
    mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.layout" icon={<span data-testid="acme-layout-icon" />} />
      </Toolbar>
    );
    expect(
      screen.getByTestId('acme-layout-icon').closest('[data-slot="slide.layout"]') !== null
    ).toBe(true);
    cleanup();
    mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.layout" asChild>
          <a href="#layout" data-testid="acme-layout" />
        </ToolbarSlot>
      </Toolbar>
    );
    const host = screen.getByTestId('acme-layout');
    expect(host.getAttribute('aria-label')).toBe('Layout');
    fireEvent.click(host);
    expect(screen.getAllByRole('menuitemradio').length).toBe(6);
    expect(warn.mock.calls.length).toBe(0);
  });

  test('a disabled Layout trigger and New slide caret stay focusable arrow stops that say why', async () => {
    const editor = await open();
    act(() => editor.setMode('view'));
    mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.add" />
        <ToolbarSlot slotId="slide.layout" />
      </Toolbar>
    );
    const caret = screen.getByRole('button', { name: 'Choose a layout' });
    const layout = screen.getByRole('button', { name: 'Layout' });
    for (const trigger of [caret, layout]) {
      expect(trigger.hasAttribute('disabled')).toBe(false);
      expect(trigger.getAttribute('aria-disabled')).toBe('true');
      const note = document.getElementById(trigger.getAttribute('aria-describedby')!);
      expect(note?.textContent).toBe('The presentation is open for viewing');
      fireEvent.click(trigger);
      expect(screen.queryAllByRole('menuitemradio').length).toBe(0);
      expect(screen.queryAllByRole('menuitem').length).toBe(0);
    }
    // The arrow keys stop on them.
    const add = screen.getByRole('button', { name: /^New slide/ });
    act(() => add.focus());
    fireEvent.keyDown(add, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(caret);
    fireEvent.keyDown(caret, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(layout);
  });

  test('asChild on Layout: a host element that is not a button opens with Enter and Space', async () => {
    const editor = await open();
    mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.layout" asChild>
          <span data-testid="acme-layout" />
        </ToolbarSlot>
      </Toolbar>
    );
    const host = screen.getByTestId('acme-layout');
    expect(host.getAttribute('role')).toBe('button');
    expect(host.tabIndex).toBe(0);
    fireEvent.keyDown(host, { key: 'Enter' });
    expect(screen.getAllByRole('menuitemradio').length).toBe(6);
    expect(host.getAttribute('aria-expanded')).toBe('true');
    fireEvent.keyDown(host, { key: ' ' });
    expect(screen.queryAllByRole('menuitemradio').length).toBe(0);
  });

  test('step 4: hidden removes a control, and its command keeps working', async () => {
    const editor = await open();
    const { container } = mount(
      editor,
      <Toolbar preset={false}>
        <ToolbarSlot slotId="slide.layout" hidden />
        <ToolbarSlot slotId="slide.add" />
      </Toolbar>
    );
    expect(container.querySelector('[data-slot="slide.layout"]') === null).toBe(true);
    const id = editor.snapshot().layouts!.find((l) => l.name === 'Blank')!.id;
    act(() => void editor.exec({ type: 'applyLayout', layoutId: id }));
    expect(layoutName(editor, 0)).toBe('Blank');
  });

  test('the gallery groups layouts under their master and moves with the arrow keys', async () => {
    const editor = await open();
    mount(editor, <LayoutGallery onChange={() => undefined} />);
    const group = screen.getByRole('group', { name: 'Fixture' });
    const tiles = within(group).getAllByRole('menuitem');
    expect(tiles.map(tileName)).toEqual([
      'Title and Content',
      'Blank',
      'Title Slide',
      'Two Content',
      'Title Only',
      'Quote',
    ]);
    // One tile takes Tab focus; arrows move by tile and by row of three.
    expect(tiles.filter((t) => t.tabIndex === 0)).toEqual([tiles[0]]);
    act(() => tiles[0]!.focus());
    fireEvent.keyDown(tiles[0]!, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(tiles[1]);
    fireEvent.keyDown(tiles[1]!, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(tiles[4]);
    fireEvent.keyDown(tiles[4]!, { key: 'Home' });
    expect(document.activeElement).toBe(tiles[0]);
    // Left from the first column leaves the gallery (a submenu then closes): the key bubbles.
    let bubbled = false;
    document.addEventListener('keydown', () => (bubbled = true), { once: true });
    fireEvent.keyDown(tiles[3]!, { key: 'ArrowLeft' });
    expect(bubbled).toBe(true);
  });
});
