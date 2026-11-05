// Sections in the rail and the slide sorter (`editor-ui-parity` 2.5): in the rail the
// sections are a tree whose headers collapse (the engine's `collapsedSections`), select their
// slides, rename in place, and carry the section menu; the sorter shows every slide in a grid
// and edits through the same commands as the rail.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen, within } from '@testing-library/react';
import {
  createPptxEditor,
  resetDevWarnings,
  runToolbarCommand,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { SlideRail } from '../panes.tsx';
import { requestSectionRename } from '../RailSections.tsx';
import { DefaultWorkspace, SlideSorter } from '../SlideSorter.tsx';
import { Toolbar, ToolbarSlot } from '../Toolbar.tsx';
import { tooltipText } from './tooltip-text.ts';

afterEach(cleanup);

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

function mount(editor: PptxEditorInstance, ui: React.ReactNode): void {
  render(<PptxEditorContext.Provider value={editor}>{ui}</PptxEditorContext.Provider>);
}

const sections = (e: PptxEditorInstance): string[] =>
  (e.snapshot().sections ?? []).map((s) => `${s.name}:${s.firstSlide}+${s.slideCount}`);
/** A section's tree item in the rail. */
const header = (name: string): HTMLElement =>
  screen.getByRole('treeitem', { name: new RegExp(`^Section ${name},`) });
const toggle = (name: string): HTMLElement =>
  header(name).querySelector<HTMLElement>('.pptx-rail__section-toggle')!;
const slide = (name: string): HTMLElement => screen.getByRole('treeitem', { name });

describe('sections in the rail', () => {
  test('are a tree: each section holds its slides, and its arrow collapses them', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    expect(screen.getByRole('tree', { name: 'Slides' }) !== null).toBe(true);
    const details = header('Details');
    expect(details.getAttribute('aria-label')).toBe('Section Details, 4 slides');
    expect(within(details).getAllByRole('treeitem').length).toBe(4);
    fireEvent.click(toggle('Details'));
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    expect(editor.snapshot().collapsedSections).toEqual([editor.snapshot().sections![1]!.id]);
    expect(within(header('Details')).queryAllByRole('treeitem')).toEqual([]);
    // Going to a slide in a collapsed section opens it.
    act(() => editor.goToSlide(3));
    expect(header('Details').getAttribute('aria-expanded')).toBe('true');
    // The collapse is a view setting: no undo step.
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('the section that holds the current slide collapses too, and Collapse all shuts all', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.click(toggle('Opening'));
    expect(header('Opening').getAttribute('aria-expanded')).toBe('false');
    // The shut section takes Tab in place of its current slide.
    expect(header('Opening').tabIndex).toBe(0);
    fireEvent.click(toggle('Opening'));
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Collapse all' }));
    expect(header('Opening').getAttribute('aria-expanded')).toBe('false');
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Expand all' }));
    expect(header('Details').getAttribute('aria-expanded')).toBe('true');
  });

  test('a click on the name selects the section’s slides, and slide keys act on them', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.click(header('Details').querySelector('.pptx-rail__section-name')!);
    expect(editor.snapshot().selectedSlides).toEqual([2, 3, 4, 5]);
    // The first slide of the section shows, not the last.
    expect(editor.snapshot().currentSlide).toBe(2);
    const focused = document.activeElement as HTMLElement;
    expect(focused.classList.contains('pptx-rail__item')).toBe(true);
    expect(header('Details').contains(focused)).toBe(true);
    fireEvent.keyDown(focused, { key: 'd', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(10);
    // A shut section keeps focus on its header.
    act(() => void editor.exec({ type: 'undo' }));
    fireEvent.click(toggle('Opening'));
    fireEvent.click(header('Opening').querySelector('.pptx-rail__section-name')!);
    expect(editor.snapshot().selectedSlides).toEqual([0, 1]);
    expect(document.activeElement === header('Opening')).toBe(true);
  });

  test('tree keys: Up from a first slide and Left go to the header; Right and Left on it', async () => {
    const editor = await open();
    act(() => editor.goToSlide(2));
    mount(editor, <SlideRail />);
    act(() => slide('Slide 3').focus());
    fireEvent.keyDown(slide('Slide 3'), { key: 'ArrowUp' });
    expect(document.activeElement === header('Details')).toBe(true);
    // A header only takes focus: the current slide stays.
    expect(editor.snapshot().currentSlide).toBe(2);
    fireEvent.keyDown(header('Details'), { key: 'ArrowLeft' });
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    fireEvent.keyDown(header('Details'), { key: 'ArrowRight' });
    expect(header('Details').getAttribute('aria-expanded')).toBe('true');
    fireEvent.keyDown(header('Details'), { key: 'ArrowDown' });
    expect(document.activeElement === slide('Slide 3')).toBe(true);
    fireEvent.keyDown(slide('Slide 3'), { key: 'ArrowDown' });
    expect(editor.snapshot().currentSlide).toBe(3);
    fireEvent.keyDown(slide('Slide 4 (hidden)'), { key: 'ArrowLeft' });
    expect(document.activeElement === header('Details')).toBe(true);
    fireEvent.keyDown(header('Details'), { key: 'Enter' });
    expect(editor.snapshot().selectedSlides).toEqual([2, 3, 4, 5]);
    expect(editor.snapshot().currentSlide).toBe(2);
  });

  test('Add section in the slide menu adds one at the slide and starts renaming it', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.contextMenu(slide('Slide 5'), { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Add section' }));
    expect(sections(editor)).toEqual(['Opening:0+2', 'Details:2+2', 'Untitled Section:4+2']);
    const field = screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement;
    expect(document.activeElement === field).toBe(true);
    fireEvent.change(field, { target: { value: 'Appendix' } });
    fireEvent.keyDown(field, { key: 'Enter' });
    expect(sections(editor)[2]).toBe('Appendix:4+2');
    // Focus goes back to the section.
    expect(document.activeElement === header('Appendix')).toBe(true);
    // Two undo steps: the rename, then the new section.
    act(() => void editor.exec({ type: 'undo' }));
    expect(sections(editor)[2]).toBe('Untitled Section:4+2');
  });

  test('double-click renames; Escape cancels, also when the field then loses focus', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.doubleClick(header('Opening').querySelector('.pptx-rail__section-name')!);
    const field = screen.getByRole('textbox', { name: 'Section name' });
    fireEvent.change(field, { target: { value: 'Start' } });
    // Editing keys stay in the field: they never act on the selected slides.
    fireEvent.keyDown(field, { key: 'Backspace' });
    fireEvent.keyDown(field, { key: 'd', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(6);
    fireEvent.keyDown(field, { key: 'Escape' });
    fireEvent.blur(field);
    expect(sections(editor)[0]).toBe('Opening:0+2');
    expect(screen.queryByRole('textbox', { name: 'Section name' }) === null).toBe(true);
    expect(document.activeElement === header('Opening')).toBe(true);
  });

  test('a rename that ends on a blur writes the name and leaves focus where it went', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.doubleClick(header('Opening').querySelector('.pptx-rail__section-name')!);
    const field = screen.getByRole('textbox', { name: 'Section name' });
    fireEvent.change(field, { target: { value: 'Start' } });
    act(() => slide('Slide 3').focus());
    expect(sections(editor)[0]).toBe('Start:0+2');
    expect(document.activeElement === slide('Slide 3')).toBe(true);
  });

  test('F2 renames and Shift+F10 opens the menu, from the section', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    act(() => header('Details').focus());
    fireEvent.keyDown(header('Details'), { key: 'F10', shiftKey: true });
    expect(screen.getByRole('menu', { name: 'Section actions' }) !== null).toBe(true);
    fireEvent.keyDown(screen.getByRole('menu'), { key: 'Escape' });
    fireEvent.keyDown(header('Details'), { key: 'F2' });
    expect(screen.getByRole('textbox', { name: 'Section name' }) !== null).toBe(true);
  });

  test('the section menu removes and moves sections', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.contextMenu(header('Opening').firstElementChild!, { clientX: 10, clientY: 10 });
    const up = screen.getByRole('menuitem', { name: 'Move section up' });
    expect(up.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(up)).toContain('The section is already at the edge');
    fireEvent.click(screen.getByRole('menuitem', { name: 'Move section down' }));
    expect(sections(editor)).toEqual(['Details:0+4', 'Opening:4+2']);
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Remove section' }));
    expect(sections(editor)).toEqual(['Opening:0+6']);
    fireEvent.contextMenu(header('Opening').firstElementChild!, { clientX: 10, clientY: 10 });
    const withSlides = screen.getByRole('menuitem', { name: 'Remove section and slides' });
    expect(withSlides.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(screen.getByRole('menuitem', { name: 'Remove all sections' }));
    expect(sections(editor)).toEqual([]);
    // Without sections the rail is a flat list again.
    expect(screen.getByRole('listbox', { name: 'Slides' }) !== null).toBe(true);
  });
});

describe('slide sorter', () => {
  const tile = (name: string): HTMLElement => screen.getByRole('treeitem', { name });

  test('shows every slide by section as a tree; a click selects, a double-click opens Normal', async () => {
    const editor = await open();
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    mount(editor, <SlideSorter />);
    expect(screen.getByRole('tree', { name: 'Slide sorter' }) !== null).toBe(true);
    const details = screen.getByRole('treeitem', { name: 'Section Details, 4 slides' });
    expect(within(details).getAllByRole('treeitem').length).toBe(4);
    expect(tile('Slide 4 (hidden)') !== null).toBe(true);
    fireEvent.click(tile('Slide 3'));
    fireEvent.click(tile('Slide 5'), { shiftKey: true });
    expect(editor.snapshot().selectedSlides).toEqual([2, 3, 4]);
    fireEvent.doubleClick(tile('Slide 6'));
    expect(editor.snapshot().presentationView).toBe('normal');
    expect(editor.snapshot().currentSlide).toBe(5);
  });

  test('shown from a menu, it focuses the current slide; focus a host put elsewhere stays', async () => {
    const editor = await open();
    act(() => editor.goToSlide(2));
    // The menu row that switched the view is gone: focus is on the body.
    (document.activeElement as HTMLElement | null)?.blur();
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    mount(editor, <SlideSorter />);
    await act(async () => void (await new Promise((r) => setTimeout(r, 50))));
    expect(document.activeElement).toBe(tile('Slide 3'));
    cleanup();
    const host = document.createElement('input');
    document.body.append(host);
    host.focus();
    mount(editor, <SlideSorter />);
    await act(async () => void (await new Promise((r) => setTimeout(r, 50))));
    expect(document.activeElement).toBe(host);
    host.remove();
  });

  test('Delete on the last tile keeps focus in the sorter, so Ctrl+Z still works', async () => {
    const editor = await open();
    mount(editor, <SlideSorter />);
    act(() => editor.goToSlide(5));
    act(() => tile('Slide 6').focus());
    fireEvent.keyDown(tile('Slide 6'), { key: 'Delete' });
    expect(editor.snapshot().slides.length).toBe(5);
    const now = document.activeElement as HTMLElement;
    expect(now.classList.contains('pptx-sorter__item')).toBe(true);
    fireEvent.keyDown(now, { key: 'z', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(6);
  });

  test('Up and Down go to the nearest tile of the row above or below, then out of the section', async () => {
    const editor = await open();
    mount(editor, <SlideSorter />);
    // Opening: one row of two tiles. Details: two rows of two tiles.
    const place = (name: string, col: number, row: number): void => {
      const rect = { left: col * 100, top: row * 80, width: 90, height: 70 };
      tile(name).getBoundingClientRect = () =>
        ({
          ...rect,
          right: rect.left + 90,
          bottom: rect.top + 70,
          x: rect.left,
          y: rect.top,
        }) as DOMRect;
    };
    place('Slide 1', 0, 0);
    place('Slide 2', 1, 0);
    place('Slide 3', 0, 1);
    place('Slide 4 (hidden)', 1, 1);
    place('Slide 5', 0, 2);
    place('Slide 6', 1, 2);
    const details = screen.getByRole('treeitem', { name: 'Section Details, 4 slides' });
    act(() => tile('Slide 4 (hidden)').focus());
    fireEvent.keyDown(tile('Slide 4 (hidden)'), { key: 'ArrowDown' });
    expect(editor.snapshot().currentSlide).toBe(5);
    expect(document.activeElement === tile('Slide 6')).toBe(true);
    fireEvent.keyDown(tile('Slide 6'), { key: 'ArrowUp' });
    expect(document.activeElement === tile('Slide 4 (hidden)')).toBe(true);
    // Past the first row of a section: its header. Down from the header: its first row.
    fireEvent.keyDown(tile('Slide 4 (hidden)'), { key: 'ArrowUp' });
    expect(document.activeElement === details).toBe(true);
    fireEvent.keyDown(details, { key: 'ArrowDown' });
    expect(document.activeElement === tile('Slide 3')).toBe(true);
    // Past the last row of a section: the row after it, the next section's header.
    act(() => tile('Slide 2').focus());
    fireEvent.keyDown(tile('Slide 2'), { key: 'ArrowDown' });
    expect(document.activeElement === details).toBe(true);
    // Ctrl+M adds a slide after the current one.
    fireEvent.keyDown(details, { key: 'ArrowDown' });
    // Ctrl+Shift+M and Ctrl+Alt+M do nothing, as on the canvas and in the rail.
    fireEvent.keyDown(tile('Slide 3'), { key: 'm', ctrlKey: true, shiftKey: true });
    fireEvent.keyDown(tile('Slide 3'), { key: 'm', ctrlKey: true, altKey: true });
    expect(editor.snapshot().slides.length).toBe(6);
    fireEvent.keyDown(tile('Slide 3'), { key: 'm', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(7);
  });

  test('keys move, select all, delete, duplicate, and redo; focus follows; a drop moves', async () => {
    const editor = await open();
    mount(editor, <SlideSorter />);
    const first = tile('Slide 1');
    fireEvent.keyDown(first, { key: 'ArrowRight' });
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(document.activeElement === tile('Slide 2')).toBe(true);
    fireEvent.keyDown(tile('Slide 2'), { key: 'End' });
    expect(editor.snapshot().currentSlide).toBe(5);
    fireEvent.keyDown(tile('Slide 6'), { key: 'a', metaKey: true });
    expect(editor.snapshot().selectedSlides).toEqual([0, 1, 2, 3, 4, 5]);
    act(() => editor.goToSlide(5));
    fireEvent.keyDown(tile('Slide 6'), { key: 'd', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(7);
    // Focus follows the copy, which is the current slide now.
    expect(document.activeElement === tile('Slide 7')).toBe(true);
    fireEvent.keyDown(tile('Slide 7'), { key: 'z', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(6);
    fireEvent.keyDown(tile('Slide 6'), { key: 'y', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(7);
    fireEvent.keyDown(tile('Slide 7'), { key: 'Delete' });
    expect(editor.snapshot().slides.length).toBe(6);
    // Drag slide 1 after slide 3: it lands third.
    fireEvent.dragStart(tile('Slide 1'), { dataTransfer: { effectAllowed: '' } });
    fireEvent.dragOver(tile('Slide 3'), { clientX: 1000 });
    fireEvent.drop(tile('Slide 3'));
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
  });

  test('section headers are in the keyboard order and take Tab when their section is shut', async () => {
    const editor = await open();
    mount(editor, <SlideSorter />);
    const opening = () => screen.getByRole('treeitem', { name: /^Section Opening,/ });
    const details = () => screen.getByRole('treeitem', { name: /^Section Details,/ });
    act(() => tile('Slide 3').focus());
    act(() => editor.goToSlide(2));
    // Up from a section's first row goes to its header; the header only takes focus.
    fireEvent.keyDown(tile('Slide 3'), { key: 'ArrowUp' });
    expect(document.activeElement === details()).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(2);
    fireEvent.keyDown(details(), { key: 'ArrowLeft' });
    fireEvent.keyDown(details(), { key: 'Delete' });
    expect(editor.snapshot().slides.length).toBe(6);
    // Every section shut: the current slide's header takes Tab.
    act(() => void editor.exec({ type: 'setSectionsCollapsed', collapsed: true }));
    expect(screen.queryAllByRole('treeitem', { name: /^Slide/ })).toEqual([]);
    expect(details().tabIndex).toBe(0);
    expect(opening().tabIndex).toBe(-1);
    fireEvent.keyDown(details(), { key: 'ArrowRight' });
    expect(details().getAttribute('aria-expanded')).toBe('true');
  });

  test('switching the view returns focus to the current slide, and refusals show', async () => {
    const editor = await open();
    mount(editor, <DefaultWorkspace normal={<p>normal view</p>} />);
    const title = editor.session!.layout(0)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [title] }));
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    expect(screen.queryByText('normal view') === null).toBe(true);
    expect(editor.snapshot().selection).toEqual({ kind: 'none' });
    expect(document.activeElement === tile('Slide 1')).toBe(true);
    // Delete every slide: refused, and the sorter says why.
    fireEvent.keyDown(tile('Slide 1'), { key: 'a', ctrlKey: true });
    fireEvent.keyDown(tile('Slide 1'), { key: 'Delete' });
    expect(screen.getByRole('status').textContent).toContain('at least one slide');
    // The zoom is the slide's zoom: in the sorter it is off and says why.
    // A disabled control stays reachable and says why (`aria-disabled`).
    const zoomIn = screen.getByRole('button', { name: 'Zoom in' }) as HTMLButtonElement;
    expect(zoomIn.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(zoomIn)).toContain('Normal view');
    // The view switch is a one-of-two choice, and an arrow key keeps focus in it.
    const sorter = screen.getByRole('radio', { name: 'Slide sorter' });
    expect(sorter.getAttribute('aria-checked')).toBe('true');
    act(() => sorter.focus());
    fireEvent.keyDown(sorter, { key: 'ArrowLeft' });
    expect(editor.snapshot().presentationView).toBe('normal');
    const normal = screen.getByRole('radio', { name: 'Normal' });
    expect(document.activeElement === normal).toBe(true);
    // A zoom menu open before a switch to the sorter stays closed on the way back.
    fireEvent.click(screen.getByRole('button', { name: /^Zoom: / }));
    expect(screen.getByRole('menu') !== null).toBe(true);
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    act(() => void editor.exec({ type: 'setPresentationView', view: 'normal' }));
    expect(screen.queryByRole('menu') === null).toBe(true);
    expect(sorter.isConnected).toBe(true);
    expect((screen.getByRole('button', { name: 'Zoom in' }) as HTMLButtonElement).disabled).toBe(
      false
    );
  });
});

describe('the section.* and view.presentationView slots', () => {
  test('take class names and state; rows that cannot run say why', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.remove" className="acme-remove" />
          <ToolbarSlot slotId="section.collapseAll" />
          <ToolbarSlot slotId="view.presentationView" className="acme-view" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const at = (id: string) => container.querySelector<HTMLElement>(`[data-slot="${id}"]`)!;
    expect(at('view.presentationView').classList.contains('acme-view')).toBe(true);
    fireEvent.click(at('view.presentationView'));
    fireEvent.click(screen.getByRole('menuitemradio', { name: 'Slide sorter' }));
    expect(editor.snapshot().presentationView).toBe('slideSorter');
    expect(at('section.remove').classList.contains('acme-remove')).toBe(true);
    act(() => void editor.exec({ type: 'removeAllSections' }));
    for (const id of ['section.remove', 'section.collapseAll']) {
      expect(at(id).hasAttribute('data-disabled')).toBe(true);
      expect(tooltipText(at(id))).toContain('This presentation has no sections');
    }
  });

  test('Add section and Rename section start the rename in the rail', async () => {
    const editor = await open();
    act(() => editor.goToSlide(4));
    mount(
      editor,
      <>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.add" />
          <ToolbarSlot slotId="section.rename" />
        </Toolbar>
        <SlideRail />
      </>
    );
    fireEvent.click(screen.getByRole('button', { name: 'Add section' }));
    expect(sections(editor)[2]).toBe('Untitled Section:4+2');
    const field = screen.getByRole('textbox', { name: 'Section name' });
    fireEvent.keyDown(field, { key: 'Escape' });
    act(() => editor.goToSlide(0));
    fireEvent.click(screen.getByRole('button', { name: 'Rename section' }));
    const again = screen.getByRole('textbox', { name: 'Section name' }) as HTMLInputElement;
    expect(again.value).toBe('Opening');
  });

  test('without a rail or a sorter, Rename section is off and says why', async () => {
    const editor = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { container } = render(
        <PptxEditorContext.Provider value={editor}>
          <Toolbar preset={false}>
            <ToolbarSlot slotId="section.rename" />
          </Toolbar>
        </PptxEditorContext.Provider>
      );
      const rename = container.querySelector<HTMLElement>('[data-slot="section.rename"]')!;
      expect(rename.hasAttribute('data-disabled')).toBe(true);
      expect(tooltipText(rename)).toContain('slide sorter to rename a section');
      // A host's own chrome hears the same: the raw run names the public way, and a request
      // that nothing takes warns.
      expect(runToolbarCommand(editor, 'section.rename')).toMatchObject({
        ok: false,
        reason: expect.stringContaining('requestSectionRename'),
      });
      requestSectionRename(editor, editor.snapshot().sections![0]!.id);
      const printed = warn.mock.calls.map(([m]) => String(m));
      expect(printed.some((m) => m.includes('runToolbarCommand("section.rename")'))).toBe(true);
      expect(printed.some((m) => m.includes('no slide rail or slide sorter'))).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });

  test('Add section without a rail or a sorter adds the section and prints no warning', async () => {
    const editor = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      mount(
        editor,
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.add" />
        </Toolbar>
      );
      act(() => editor.goToSlide(4));
      fireEvent.click(screen.getByRole('button', { name: 'Add section' }));
      expect(sections(editor)).toHaveLength(3);
      expect(warn.mock.calls.length).toBe(0);
    } finally {
      warn.mockRestore();
    }
  });

  test('hidden removes a section control, and its command keeps working', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="section.removeAll" hidden />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    expect(container.querySelector('[data-slot="section.removeAll"]') === null).toBe(true);
    act(() => void editor.exec({ type: 'removeAllSections' }));
    expect(sections(editor)).toEqual([]);
  });
});

describe('keys on a focused section header in the rail', () => {
  test('do not act on the selected slides; ArrowRight goes in and selects', async () => {
    const editor = await open();
    act(() => editor.goToSlide(2));
    mount(editor, <SlideRail />);
    act(() => slide('Slide 3').focus());
    fireEvent.keyDown(slide('Slide 3'), { key: 'ArrowUp' });
    expect(document.activeElement === header('Details')).toBe(true);
    for (const key of [{ key: 'Delete' }, { key: 'Backspace' }, { key: 'd', ctrlKey: true }])
      fireEvent.keyDown(header('Details'), key);
    expect(editor.snapshot().slides.length).toBe(6);
    act(() => editor.goToSlide(0));
    fireEvent.keyDown(header('Details'), { key: 'ArrowRight' });
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(document.activeElement === slide('Slide 3')).toBe(true);
  });
});
