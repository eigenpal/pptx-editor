import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, notesText, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { EditNotice } from '../EditNotice.tsx';
import { NotesPane, SlideRail } from '../panes.tsx';

// E1: the rail buttons keep focus in the notes field, so a draft typed on
// one slide must land on that slide, never on the slide that is current after the command.

async function setup(): Promise<{ editor: PptxEditorInstance; field: HTMLElement }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(
    buildPptx({
      slides: [
        { shapes: '', notes: 'one' },
        { shapes: '', notes: 'two' },
      ],
    })
  );
  expect(r.ok).toBe(true);
  render(
    <PptxEditorContext.Provider value={editor}>
      <SlideRail />
      <NotesPane />
      <EditNotice />
    </PptxEditorContext.Provider>
  );
  const field = screen.getByRole('textbox', { name: 'Speaker notes' });
  return { editor, field };
}

const notes = (e: PptxEditorInstance): string[] =>
  e.snapshot().slides.map((_, i) => notesText(e.session!.store.package, i));

/** Type into the notes field: the browser edits the DOM, then fires `input`. */
function type(field: HTMLElement, text: string): void {
  field.focus();
  field.textContent = text;
  fireEvent.input(field);
}

/** Click a rail button; chrome keeps the focus where it is. */
function click(name: string, pointer = false): void {
  const button = screen.getByRole('button', { name });
  if (pointer) fireEvent.pointerDown(button);
  fireEvent.mouseDown(button);
  fireEvent.click(button);
}

afterEach(cleanup);

describe('NotesPane while the notes field has focus', () => {
  test('New slide writes the draft to the slide it was typed on', async () => {
    const { editor, field } = await setup();
    type(field, 'one edited');
    click('New slide');
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(field.textContent).toBe('');
    fireEvent.blur(field);
    expect(notes(editor)).toEqual(['one edited', '', 'two']);
  });

  test('Delete slide never writes the draft onto the next slide', async () => {
    const { editor, field } = await setup();
    type(field, 'one edited');
    click('Delete slide');
    fireEvent.blur(field);
    expect(notes(editor)).toEqual(['two']);
    expect(field.textContent).toBe('two');
  });

  test('Delete slide after a pointer press keeps the draft in the history', async () => {
    const { editor, field } = await setup();
    type(field, 'one edited');
    click('Delete slide', true);
    expect(notes(editor)).toEqual(['two']);
    editor.exec({ type: 'undo' });
    expect(notes(editor)).toEqual(['one edited', 'two']);
  });

  test('Move slide keeps the draft with its slide', async () => {
    const { editor, field } = await setup();
    type(field, 'one edited');
    click('Move slide down');
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(field.textContent).toBe('one edited');
    fireEvent.blur(field);
    expect(notes(editor)).toEqual(['two', 'one edited']);
  });

  test('two edits are two undo steps', async () => {
    const { editor, field } = await setup();
    type(field, 'first');
    fireEvent.blur(field);
    type(field, 'second');
    fireEvent.blur(field);
    editor.exec({ type: 'undo' });
    expect(notes(editor)[0]).toBe('first');
  });

  test('a refused edit is shown', async () => {
    const { editor } = await setup();
    act(() => void editor.input('text', 'x'));
    expect(screen.getByRole('status').textContent).toContain('Select a shape that has text');
  });
});
