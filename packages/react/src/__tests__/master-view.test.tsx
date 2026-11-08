// The slide master view in the React adapter: the slide rail lists masters and layouts, the
// content shows the bar that names the template and closes the view, and the notes pane hides.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { NotesPane } from '../NotesPane.tsx';
import { SlideRail } from '../panes.tsx';
import { PptxEditorContent } from '../PptxEditorRoot.tsx';

afterEach(cleanup);

async function mounted(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  render(
    <PptxEditorContext.Provider value={editor}>
      <SlideRail />
      <PptxEditorContent />
      <NotesPane />
    </PptxEditorContext.Provider>
  );
  return editor;
}

const openMaster = (editor: PptxEditorInstance): void => {
  act(() => {
    expect(editor.exec({ type: 'setPresentationView', view: 'slideMaster' }).ok).toBe(true);
  });
};

describe('the slide master view in React', () => {
  test('the rail lists the master and its layouts with their use', async () => {
    const editor = await mounted();
    openMaster(editor);
    const rail = screen.getByRole('listbox', { name: 'Masters and layouts' });
    const items = rail.querySelectorAll<HTMLElement>('[role="option"]');
    expect(items.length).toBe(editor.snapshot().slides.length);
    expect(items[0]!.dataset.template).toBe('master');
    expect(items[0]!.querySelector('.pptx-master-rail__name')!.textContent).toMatch(/^Master/);
    const current = rail.querySelector<HTMLElement>('[aria-current="true"]')!;
    expect(Number(current.dataset.index)).toBe(editor.snapshot().currentSlide);
    expect(current.textContent).toMatch(/, Used by (1 slide|\d+ slides)$/);
    expect(current.querySelector('.pptx-master-rail__usage')!.textContent).toMatch(/^Used by/);
  });

  test('a click and the arrow keys show another template; Home goes to the master', async () => {
    const editor = await mounted();
    openMaster(editor);
    const items = () => screen.getAllByRole('option');
    act(() => fireEvent.click(items()[2]!));
    expect(editor.snapshot().currentSlide).toBe(Number(items()[2]!.dataset.index));
    act(() => fireEvent.keyDown(items()[2]!, { key: 'ArrowDown' }));
    expect(editor.snapshot().currentSlide).toBe(Number(items()[3]!.dataset.index));
    act(() => fireEvent.keyDown(items()[3]!, { key: 'Home' }));
    expect(editor.snapshot().currentSlide).toBe(Number(items()[0]!.dataset.index));
    expect(screen.getByRole('region', { name: /^Master: / }).textContent).toContain(
      'every layout of this master'
    );
  });

  test('the bar closes the view; the notes pane hides while it is open', async () => {
    const editor = await mounted();
    act(() => editor.goToSlide(2));
    expect(document.querySelector('.pptx-notes')).not.toBeNull();
    openMaster(editor);
    expect(document.querySelector('.pptx-notes')).toBeNull();
    expect(screen.queryByRole('listbox', { name: 'Masters and layouts' })).not.toBeNull();
    act(() => fireEvent.click(screen.getByRole('button', { name: 'Close master' })));
    expect(editor.snapshot().presentationView).toBe('normal');
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(screen.queryByRole('region', { name: /^(Master|Layout): / })).toBeNull();
    // Focus goes to the slide, not to the page, as the bar leaves.
    expect(document.activeElement?.closest('.pptx-canvas')).not.toBeNull();
    expect(document.querySelector('.pptx-notes')).not.toBeNull();
    // The slide rail is back, on the slide the view opened from.
    const back = document.querySelector<HTMLElement>('.pptx-rail__item[aria-current="true"]');
    expect(back).not.toBeNull();
    expect(back!.classList.contains('pptx-master-rail__item')).toBe(false);
  });
});
