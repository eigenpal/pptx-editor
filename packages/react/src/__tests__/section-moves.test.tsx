// Moves across a section header in the rail and the sorter:
// a drop before a section's first slide (below its header) puts the slide into that section,
// and Ctrl/Cmd+Down from the last slide of a section makes it the first slide of the next.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, createEvent, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, slideStep, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { SlideRail } from '../panes.tsx';
import { SlideSorter } from '../SlideSorter.tsx';

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
const item = (name: string): HTMLElement => screen.getByRole('treeitem', { name });

/** Drag `from` over `to`, on the side where the pointer is at `at` (in px along the list). */
function dragOver(to: HTMLElement, axis: 'clientX' | 'clientY', at: number): void {
  const e = createEvent.dragOver(to);
  Object.defineProperty(e, axis, { value: at });
  fireEvent(to, e);
}

describe('moves across a section header', () => {
  test('slideStep changes the section at a header, and names the section inside one', () => {
    const grouped = [
      { id: 'a', name: 'A', firstSlide: 0, slideCount: 2 },
      { id: 'b', name: 'B', firstSlide: 2, slideCount: 4 },
    ];
    expect(slideStep([1], 6, 1, grouped)).toEqual({ to: 1, section: 'b' });
    expect(slideStep([2], 6, -1, grouped)).toEqual({ to: 2, section: 'a' });
    expect(slideStep([3], 6, -1, grouped)).toEqual({ to: 2, section: 'b' });
    expect(slideStep([0], 6, 1, grouped)).toEqual({ to: 1, section: 'a' });
    expect(slideStep([0], 6, -1, grouped)).toBeNull();
    expect(slideStep([5], 6, 1, grouped)).toBeNull();
    expect(slideStep([1], 6, 1, null)).toEqual({ to: 2 });
  });

  test('rail: a drop before the first slide of a section joins it; after the last slide stays', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    fireEvent.dragStart(item('Slide 1'), { dataTransfer: { effectAllowed: '' } });
    dragOver(item('Slide 3'), 'clientY', -1);
    fireEvent.drop(item('Slide 3'));
    expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
    expect(editor.snapshot().currentSlide).toBe(1);
    act(() => void editor.exec({ type: 'undo' }));
    // The same place, dropped after the last slide of Opening (above the header).
    fireEvent.dragStart(item('Slide 1'), { dataTransfer: { effectAllowed: '' } });
    dragOver(item('Slide 2'), 'clientY', 1000);
    fireEvent.drop(item('Slide 2'));
    expect(sections(editor)).toEqual(['Opening:0+2', 'Details:2+4']);
    expect(editor.snapshot().currentSlide).toBe(1);
  });

  test('rail: Ctrl+Down from the last slide of a section makes it the first of the next', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    act(() => editor.goToSlide(1));
    const before = editor.snapshot().slides.map((s) => s.id);
    fireEvent.keyDown(item('Slide 2'), { key: 'ArrowDown', ctrlKey: true });
    expect(editor.snapshot().slides.map((s) => s.id)).toEqual(before);
    expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
    // The next step moves it inside the section.
    fireEvent.keyDown(item('Slide 2'), { key: 'ArrowDown', ctrlKey: true });
    expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
    expect(editor.snapshot().currentSlide).toBe(2);
    // Ctrl+Up back across the header: it is the last slide of Opening again, in place.
    fireEvent.keyDown(item('Slide 3'), { key: 'ArrowUp', ctrlKey: true });
    fireEvent.keyDown(item('Slide 2'), { key: 'ArrowUp', ctrlKey: true });
    expect(sections(editor)).toEqual(['Opening:0+2', 'Details:2+4']);
    expect(editor.snapshot().slides.map((s) => s.id)).toEqual(before);
  });

  test('rail and sorter: a step into a collapsed section opens it, and focus stays on the slide', async () => {
    for (const sorter of [false, true]) {
      const editor = await open();
      const details = editor.snapshot().sections![1]!.id;
      const order = editor.snapshot().slides.map((s) => s.id);
      if (sorter) act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
      act(() => editor.goToSlide(1));
      act(
        () =>
          void editor.exec({ type: 'setSectionsCollapsed', collapsed: true, sectionIds: [details] })
      );
      mount(editor, sorter ? <SlideSorter /> : <SlideRail />);
      const key = sorter ? 'ArrowRight' : 'ArrowDown';
      act(() => item('Slide 2').focus());
      fireEvent.keyDown(item('Slide 2'), { key, ctrlKey: true });
      expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
      expect(editor.snapshot().collapsedSections ?? []).not.toContain(details);
      const header = screen.getByRole('treeitem', { name: /^Section Details,/ });
      expect(header.getAttribute('aria-expanded')).toBe('true');
      expect(document.activeElement === item('Slide 2')).toBe(true);
      // Opening the section is a view setting: the only undo step is the move.
      expect(editor.snapshot().canUndo).toBe(true);
      act(() => void editor.exec({ type: 'undo' }));
      expect(editor.snapshot().canUndo).toBe(false);
      expect(editor.snapshot().slides.map((s) => s.id)).toEqual(order);
      expect(sections(editor)).toEqual(['Opening:0+2', 'Details:2+4']);
      act(() => void editor.exec({ type: 'redo' }));
      expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
      act(() => item('Slide 2').focus());
      // A further step works: the slide moves inside Details.
      fireEvent.keyDown(item('Slide 2'), { key, ctrlKey: true });
      expect(editor.snapshot().currentSlide).toBe(2);
      expect(document.activeElement === item('Slide 3')).toBe(true);
      cleanup();
    }
  });

  test('rail: Ctrl+Up inside a section keeps the slide in it', async () => {
    const editor = await open();
    mount(editor, <SlideRail />);
    act(() => editor.goToSlide(3));
    fireEvent.keyDown(item('Slide 4 (hidden)'), { key: 'ArrowUp', ctrlKey: true });
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(sections(editor)).toEqual(['Opening:0+2', 'Details:2+4']);
  });

  test('sorter: Cmd+Right from the last slide of a section makes it the first of the next', async () => {
    const editor = await open();
    act(() => void editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
    mount(editor, <SlideSorter />);
    act(() => editor.goToSlide(1));
    fireEvent.keyDown(item('Slide 2'), { key: 'ArrowRight', metaKey: true });
    expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
    // A drop before a section's first tile joins that section.
    act(() => void editor.exec({ type: 'undo' }));
    fireEvent.dragStart(item('Slide 1'), { dataTransfer: { effectAllowed: '' } });
    dragOver(item('Slide 3'), 'clientX', -1);
    fireEvent.drop(item('Slide 3'));
    expect(sections(editor)).toEqual(['Opening:0+1', 'Details:1+5']);
    expect(editor.snapshot().currentSlide).toBe(1);
  });
});
