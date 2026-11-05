// Add Section and Rename Section from the toolbar parts (with `icon` and `asChild`), `Menu.Item`,
// and `useEditorCommand`, in a deck with sections and in one without. Context menu rows run
// through the same `runChromeRow` as the menu bar.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext, useEditorCommand, type UseEditorCommandReturn } from '../context.ts';
import { PptxEditor } from '../index.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { SlideRail, ZoomPill } from '../panes.tsx';
import { SlideSorter, ViewSwitch } from '../SlideSorter.tsx';
import { Toolbar, ToolbarSlot } from '../Toolbar.tsx';

afterEach(cleanup);

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const plain = (): Uint8Array =>
  buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] });

function mount(editor: PptxEditorInstance, ui: React.ReactNode): void {
  render(<PptxEditorContext.Provider value={editor}>{ui}</PptxEditorContext.Provider>);
}

const addSection = (
  <Toolbar preset={false}>
    <ToolbarSlot slotId="section.add" />
  </Toolbar>
);

describe('Add section in a deck without sections', () => {
  test('opens the name field of the new section in the rail', async () => {
    const editor = await open(plain());
    expect(editor.snapshot().sections ?? []).toEqual([]);
    mount(
      editor,
      <>
        {addSection}
        <SlideRail />
      </>
    );
    act(() => editor.goToSlide(1));
    fireEvent.click(screen.getByRole('button', { name: 'Add section' }));
    expect(editor.snapshot().sections!.map((s) => [s.name, s.firstSlide])).toEqual([
      ['Default Section', 0],
      ['Untitled Section', 1],
    ]);
    const field = screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement;
    expect(document.activeElement === field).toBe(true);
    fireEvent.change(field, { target: { value: 'Middle' } });
    fireEvent.keyDown(field, { key: 'Enter' });
    // The section at the current slide takes the typed name; "Default Section" keeps its name.
    expect(editor.snapshot().sections!.map((s) => [s.name, s.firstSlide])).toEqual([
      ['Default Section', 0],
      ['Middle', 1],
    ]);
  });

  test('opens the name field of the new section in the sorter', async () => {
    const editor = await open(plain());
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    mount(
      editor,
      <>
        {addSection}
        <SlideSorter />
      </>
    );
    fireEvent.click(screen.getByRole('button', { name: 'Add section' }));
    const field = screen.getByRole('textbox', { name: 'Section name' });
    expect(document.activeElement === field).toBe(true);
  });

  test('a closed rail does not take the request', async () => {
    const editor = await open(plain());
    mount(
      editor,
      <>
        {addSection}
        <SlideRail open={false} />
      </>
    );
    fireEvent.click(screen.getByRole('button', { name: 'Add section' }));
    expect(editor.snapshot().sections!.length).toBeGreaterThan(0);
    expect(screen.queryByRole('textbox', { name: 'Section name' }) === null).toBe(true);
  });
});

describe('section.rename and section.add outside the toolbar', () => {
  let rename: UseEditorCommandReturn | null = null;
  let add: UseEditorCommandReturn | null = null;
  function Probe() {
    rename = useEditorCommand('section.rename');
    add = useEditorCommand('section.add');
    return null;
  }

  test('useEditorCommand: Rename is off without a rail and says why; with one, it opens the field', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    const view = render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(rename!.isEnabled).toBe(false);
    expect(rename!.disabledReason).toBe('disabled.noSectionPane');
    expect(rename!.execute()).toBe(false);
    view.rerender(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
        <SlideRail />
      </PptxEditorContext.Provider>
    );
    expect(rename!.isEnabled).toBe(true);
    act(() => void rename!.execute());
    const field = screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement;
    expect(field.value).toBe('Opening');
    fireEvent.keyDown(field, { key: 'Escape' });
    // Add Section adds a section at the current slide and asks for its name.
    act(() => editor.goToSlide(4));
    act(() => void add!.execute());
    expect(editor.snapshot().sections!.map((x) => x.name)).toContain('Untitled Section');
    expect((screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement).value).toBe(
      'Untitled Section'
    );
  });

  test('Menu.Item runs Rename section and Add section with the name field', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.Menu preset={false}>
          <PptxEditor.Menu.Edit>
            <PptxEditor.Menu.Item slotId="section.rename" />
            <PptxEditor.Menu.Item slotId="section.add" />
          </PptxEditor.Menu.Edit>
        </PptxEditor.Menu>
        <SlideRail />
      </PptxEditorContext.Provider>
    );
    const openEdit = (): HTMLElement => {
      act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Edit' })));
      return document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
    };
    let panel = openEdit();
    act(() => void fireEvent.click(panel.querySelector('[data-slot="section.rename"]')!));
    const field = screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement;
    expect(field.value).toBe('Opening');
    fireEvent.keyDown(field, { key: 'Escape' });
    act(() => editor.goToSlide(4));
    panel = openEdit();
    act(() => void fireEvent.click(panel.querySelector('[data-slot="section.add"]')!));
    expect((screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement).value).toBe(
      'Untitled Section'
    );
  });
});

describe('ContextMenu.Slot section.rename', () => {
  test('a host row on the canvas menu opens the name field in the rail', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditorRoot
        onReady={(e) => (editor = e)}
        popups={{
          contextMenu: (props) => (
            <PptxEditor.ContextMenu {...props}>
              <PptxEditor.ContextMenu.Slot slotId="section.rename" />
            </PptxEditor.ContextMenu>
          ),
        }}
      >
        <PptxEditorContent />
        <SlideRail />
      </PptxEditorRoot>
    );
    await act(async () => void (await editor!.load(buildSlideLayoutsDeck())));
    const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
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
    act(() => void fireEvent.contextMenu(stage, { clientX: 1270, clientY: 710 }));
    const row = document.querySelector<HTMLElement>(
      '.pptx-contextmenu [data-slot="section.rename"]'
    )!;
    expect(row.getAttribute('aria-disabled')).not.toBe('true');
    act(() => void fireEvent.click(row));
    const field = screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement;
    expect(field.value).toBe('Opening');
  });
});

describe('the section.add and section.rename toolbar parts', () => {
  test('take an icon and asChild, and keep their behavior', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    mount(
      editor,
      <>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.add" icon={<b data-testid="add-icon" />} />
          <ToolbarSlot slotId="section.rename" asChild>
            <button type="button" data-testid="rename" />
          </ToolbarSlot>
        </Toolbar>
        <SlideRail />
      </>
    );
    const add = document.querySelector<HTMLElement>('[data-slot="section.add"]')!;
    expect(add.querySelector('[data-testid="add-icon"]') !== null).toBe(true);
    const rename = screen.getByTestId('rename');
    expect(rename.getAttribute('data-slot')).toBe('section.rename');
    expect(rename.getAttribute('aria-label')).toContain('Rename section');
    fireEvent.click(rename);
    expect((screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement).value).toBe(
      'Opening'
    );
    fireEvent.keyDown(screen.getByRole('textbox', { name: 'Section name' }), { key: 'Escape' });
    act(() => editor.goToSlide(4));
    fireEvent.click(add);
    expect((screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement).value).toBe(
      'Untitled Section'
    );
  });
});

describe('Collapse all and focus', () => {
  const header = (name: string): HTMLElement =>
    screen.getByRole('treeitem', { name: new RegExp(`^Section ${name},`) });

  test('the toolbar part moves focus from the hidden slide to its section header', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    act(() => editor.goToSlide(3));
    mount(
      editor,
      <>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.collapseAll" />
        </Toolbar>
        <SlideRail />
      </>
    );
    act(() => screen.getByRole('treeitem', { name: 'Slide 4 (hidden)' }).focus());
    // Chrome keeps focus where it is: the click does not move it to the button.
    const button = document.querySelector<HTMLElement>('[data-slot="section.collapseAll"]')!;
    fireEvent.mouseDown(button);
    fireEvent.click(button);
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement === header('Details')).toBe(true);
  });

  test('the section menu row moves focus to the header of the current section', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    act(() => editor.goToSlide(1));
    mount(editor, <SlideRail />);
    act(() => screen.getByRole('treeitem', { name: 'Slide 2' }).focus());
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Collapse all' }));
    expect(header('Opening').getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement === header('Opening')).toBe(true);
  });

  test('in the sorter too', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    act(() => editor.goToSlide(3));
    mount(
      editor,
      <>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.collapseAll" />
        </Toolbar>
        <SlideSorter />
      </>
    );
    act(() => screen.getByRole('treeitem', { name: 'Slide 4 (hidden)' }).focus());
    fireEvent.click(document.querySelector<HTMLElement>('[data-slot="section.collapseAll"]')!);
    expect(document.activeElement === header('Details')).toBe(true);
  });
});

describe('Move section up and down and focus', () => {
  const header = (name: string): HTMLElement =>
    screen.getByRole('treeitem', { name: new RegExp(`^Section ${name},`) });

  test('focus stays on the header of the section that moved', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    mount(editor, <SlideRail />);
    act(() => header('Details').focus());
    fireEvent.keyDown(header('Details'), { key: 'F10', shiftKey: true });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Move section up' }));
    expect(editor.snapshot().sections!.map((s) => s.name)).toEqual(['Details', 'Opening']);
    expect(document.activeElement === header('Details')).toBe(true);
    fireEvent.keyDown(header('Details'), { key: 'F10', shiftKey: true });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Move section down' }));
    expect(editor.snapshot().sections!.map((s) => s.name)).toEqual(['Opening', 'Details']);
    expect(document.activeElement === header('Details')).toBe(true);
  });
});

describe('a section header that lost focus', () => {
  test('does not take focus back when its section moves later', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    mount(editor, <SlideRail />);
    const details = screen.getByRole('treeitem', { name: /^Section Details,/ });
    act(() => details.focus());
    // Focus goes to nothing: a click on an empty part of the page.
    act(() => details.blur());
    expect(document.activeElement === document.body).toBe(true);
    const id = editor.snapshot().sections![1]!.id;
    act(() => void editor.exec({ type: 'moveSection', sectionId: id, direction: -1 }));
    expect(editor.snapshot().sections!.map((s) => s.name)).toEqual(['Details', 'Opening']);
    expect(document.activeElement === document.body).toBe(true);
  });
});

describe('the zoom level and the view switch when they are off', () => {
  test('say why through aria-describedby, and a click does nothing', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    mount(editor, <ZoomPill />);
    const level = document.querySelector<HTMLElement>('.pptx-zoom__level')!;
    expect(level.getAttribute('aria-disabled')).toBe('true');
    const note = document.getElementById(level.getAttribute('aria-describedby')!)!;
    expect(note.textContent!.length).toBeGreaterThan(0);
    fireEvent.click(level);
    expect(screen.queryByRole('menu') === null).toBe(true);
    cleanup();
    // Without an editor, the view switch is off.
    render(
      <PptxEditorContext.Provider value={null}>
        <ViewSwitch />
      </PptxEditorContext.Provider>
    );
    for (const radio of screen.getAllByRole('radio')) {
      expect(radio.getAttribute('aria-disabled')).toBe('true');
      const said = document.getElementById(radio.getAttribute('aria-describedby')!)!;
      expect(said.textContent!.length).toBeGreaterThan(0);
    }
  });
});
