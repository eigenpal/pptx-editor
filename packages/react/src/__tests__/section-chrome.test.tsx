// The section and sorter chrome after review (PR eigenpal/pptx-editor-miro#39): the `railContextMenu` and
// `sectionContextMenu` popup entries in the rail and the sorter, rename registration, focus
// after a collapse and a removal, a collapsed section that keeps shut when selected by name,
// sorter keys, the Slide Sorter warning, and controls that stay focusable while off.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { SlideRail, ZoomPill } from '../panes.tsx';
import { DialogProvider, type PptxEditorPopups } from '../popups.tsx';
import { requestSectionRename } from '../RailSections.tsx';
import { SlideSorter, SlideSorterWatch, ViewSwitch } from '../SlideSorter.tsx';

afterEach(cleanup);

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

const tree = (
  editor: PptxEditorInstance | null,
  ui: React.ReactNode,
  popups?: PptxEditorPopups
) => (
  <PptxEditorContext.Provider value={editor}>
    <DialogProvider {...(popups ? { popups } : {})}>{ui}</DialogProvider>
  </PptxEditorContext.Provider>
);

const header = (name: string): HTMLElement =>
  screen.getByRole('treeitem', { name: new RegExp(`^Section ${name},`) });
const slide = (name: string): HTMLElement => screen.getByRole('treeitem', { name });
const sections = (e: PptxEditorInstance): string[] =>
  (e.snapshot().sections ?? []).map((s) => `${s.name}:${s.firstSlide}+${s.slideCount}`);

describe('popup entries of the rail and the sorter', () => {
  test('railContextMenu false shows no menu in the sorter; a function renders the host menu', async () => {
    const editor = await open();
    const view = render(tree(editor, <SlideSorter />, { railContextMenu: false }));
    fireEvent.contextMenu(slide('Slide 3'), { clientX: 10, clientY: 10 });
    expect(screen.queryByRole('menu') === null).toBe(true);
    view.rerender(
      tree(editor, <SlideSorter />, {
        railContextMenu: ({ menu }) => <div role="menu" aria-label={`Host ${menu.index}`} />,
      })
    );
    fireEvent.contextMenu(slide('Slide 3'), { clientX: 10, clientY: 10 });
    expect(screen.getByRole('menu', { name: 'Host 2' }) !== null).toBe(true);
  });

  test('sectionContextMenu false shows no section menu; a function renders the host menu', async () => {
    const editor = await open();
    const view = render(tree(editor, <SlideRail />, { sectionContextMenu: false }));
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    expect(screen.queryByRole('menu') === null).toBe(true);
    let renamed = false;
    view.rerender(
      tree(editor, <SlideRail />, {
        sectionContextMenu: ({ section, onRename }) => (
          <div role="menu" aria-label={`Host ${section.name}`}>
            <button type="button" onClick={() => ((renamed = true), onRename())}>
              Rename
            </button>
          </div>
        ),
      })
    );
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    expect(screen.getByRole('menu', { name: 'Host Details' }) !== null).toBe(true);
    fireEvent.click(screen.getByRole('button', { name: 'Rename' }));
    expect(renamed).toBe(true);
    expect(screen.getByRole('textbox', { name: 'Section name' }) !== null).toBe(true);
  });

  test('the sectionContextMenu entry covers the sorter too', async () => {
    const editor = await open();
    render(tree(editor, <SlideSorter />, { sectionContextMenu: false }));
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    expect(screen.queryByRole('menu') === null).toBe(true);
  });
});

describe('section rename registration', () => {
  test('a closed rail does not rename in place, and closing it ends a rename', async () => {
    const editor = await open();
    const sectionId = editor.snapshot().sections![0]!.id;
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const view = render(tree(editor, <SlideRail open={false} />));
      act(() => requestSectionRename(editor, sectionId));
      expect(warn.mock.calls.some((c) => String(c[0]).includes('requestSectionRename'))).toBe(true);
      view.rerender(tree(editor, <SlideRail open />));
      act(() => requestSectionRename(editor, sectionId));
      expect(screen.getByRole('textbox', { name: 'Section name' }) !== null).toBe(true);
      view.rerender(tree(editor, <SlideRail open={false} />));
      view.rerender(tree(editor, <SlideRail open />));
      expect(screen.queryByRole('textbox', { name: 'Section name' }) === null).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });

  test('an open rail without sections takes the request; a closed one does not', async () => {
    const editor = await open();
    act(() => void editor.exec({ type: 'removeAllSections' }));
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      // Add Section in a flat deck asks for the name before the new header renders.
      const view = render(tree(editor, <SlideRail />));
      act(() => requestSectionRename(editor, 'x'));
      expect(warn.mock.calls.some((c) => String(c[0]).includes('requestSectionRename'))).toBe(
        false
      );
      view.rerender(tree(editor, <SlideRail open={false} />));
      act(() => requestSectionRename(editor, 'x'));
      expect(warn.mock.calls.some((c) => String(c[0]).includes('requestSectionRename'))).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });
});

describe('sections and focus', () => {
  test('a click on the name of a collapsed section selects its slides and keeps it collapsed', async () => {
    const editor = await open();
    render(tree(editor, <SlideRail />));
    fireEvent.click(header('Details').querySelector('.pptx-rail__section-toggle')!);
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    fireEvent.click(header('Details').querySelector('.pptx-rail__section-name')!);
    expect(editor.snapshot().selectedSlides).toEqual([2, 3, 4, 5]);
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    // Going to a slide of the section in another way still opens it.
    act(() => editor.goToSlide(4));
    expect(header('Details').getAttribute('aria-expanded')).toBe('true');
  });

  test('collapsing the section of the focused slide moves focus to its header', async () => {
    const editor = await open();
    act(() => editor.goToSlide(2));
    render(tree(editor, <SlideRail />));
    act(() => slide('Slide 3').focus());
    fireEvent.click(header('Details').querySelector('.pptx-rail__section-toggle')!);
    expect(header('Details').getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement === header('Details')).toBe(true);
  });

  test('after Remove all sections, focus is on the current slide', async () => {
    const editor = await open();
    act(() => editor.goToSlide(3));
    render(tree(editor, <SlideRail />));
    fireEvent.contextMenu(header('Details').firstElementChild!, { clientX: 10, clientY: 10 });
    fireEvent.click(screen.getByRole('menuitem', { name: 'Remove all sections' }));
    expect(sections(editor)).toEqual([]);
    const focused = document.activeElement as HTMLElement;
    expect(focused.getAttribute('aria-current')).toBe('true');
    expect(focused.dataset.index).toBe('3');
  });
});

describe('sorter keys and the view', () => {
  test('Cmd/Ctrl+Arrow moves the selected slides one step in grid order', async () => {
    const editor = await open();
    render(tree(editor, <SlideSorter />));
    fireEvent.click(slide('Slide 3'));
    fireEvent.keyDown(slide('Slide 3'), { key: 'ArrowRight', ctrlKey: true });
    expect(editor.snapshot().selectedSlides).toEqual([3]);
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowUp', metaKey: true });
    // From the first slide of Details, the step crosses the header: the slide stays third and
    // becomes the last slide of Opening.
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowLeft', metaKey: true });
    expect(editor.snapshot().selectedSlides).toEqual([2]);
    expect(editor.snapshot().sections!.map((s) => s.slideCount)).toEqual([3, 3]);
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowLeft', metaKey: true });
    expect(editor.snapshot().selectedSlides).toEqual([1]);
    // Each move is one undo step.
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().selectedSlides).toEqual([2]);
    expect(editor.snapshot().sections!.map((s) => s.slideCount)).toEqual([3, 3]);
  });

  test('Slide Sorter with no sorter mounted prints a development warning', async () => {
    const editor = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const view = render(tree(editor, <SlideSorterWatch />));
      act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
      expect(warn.mock.calls.some((c) => String(c[0]).includes('Slide Sorter'))).toBe(true);
      act(() => void editor.exec({ type: 'setPresentationView', view: 'normal' }));
      warn.mockClear();
      resetDevWarnings();
      view.rerender(
        tree(
          editor,
          <>
            <SlideSorter />
            <SlideSorterWatch />
          </>
        )
      );
      act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
      expect(warn.mock.calls.some((c) => String(c[0]).includes('Slide Sorter'))).toBe(false);
    } finally {
      warn.mockRestore();
    }
  });

  test('the zoom level and the view switch stay focusable while off', async () => {
    const editor = await open();
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    render(tree(editor, <ZoomPill />));
    // The sorter draws no slide at the zoom: the level names no percentage there.
    const level = screen.getByRole('button', { name: 'Zoom' });
    expect(level.textContent).not.toContain('%');
    expect(level.hasAttribute('disabled')).toBe(false);
    expect(level.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(level);
    expect(screen.queryByRole('menu') === null).toBe(true);
    cleanup();
    render(tree(null, <ViewSwitch />));
    for (const radio of screen.getAllByRole('radio'))
      expect(radio.getAttribute('aria-disabled')).toBe('true');
  });
});
