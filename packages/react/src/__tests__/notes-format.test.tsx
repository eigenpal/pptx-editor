// Speaker notes formatting in the notes pane (`parity-baseline` 4.7): the notes field shows each
// run's formatting while you read and while you type, and Ctrl/Cmd+B, I, U or the pane's
// buttons format the selected text as one undo step. Ctrl/Cmd+Z in the field undoes deck edits.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  createPptxEditor,
  notesFieldText,
  notesFormatsAt,
  notesSelection,
  selectNotes,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext, useEditorCommand } from '../context.ts';
import { NotesPane, SlideRail } from '../panes.tsx';
import { Toolbar, ToolbarSlot } from '../Toolbar.tsx';

afterEach(cleanup);

async function setup(slide = 1): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  editor.goToSlide(slide);
  render(
    <PptxEditorContext.Provider value={editor}>
      <NotesPane />
    </PptxEditorContext.Provider>
  );
  return editor;
}

const field = (): HTMLElement => screen.getByRole('textbox', { name: 'Speaker notes' });
const runs = (e: PptxEditorInstance, p: number): string[] =>
  e.snapshot().notesFormatted![p]!.runs.map((r) => `${r.text}${r.bold ? '*' : ''}`);
const FIRST = 'Open with the headline number and pause.';
const button = (name: string): HTMLElement => screen.getByRole('button', { name });

/** Focus the field and select `[start, end)` of its plain text. */
function select(start: number, end = start): void {
  act(() => {
    field().focus();
    selectNotes(field(), start, end);
  });
  fireEvent.keyUp(field(), { key: 'Shift' });
}

/**
 * Type `text` at the caret the way the browser does: `beforeinput`, which the field can take
 * over, else into the text node, then `input`.
 */
function typeAtCaret(text: string): void {
  const before = new InputEvent('beforeinput', {
    inputType: 'insertText',
    data: text,
    bubbles: true,
    cancelable: true,
  });
  act(() => void field().dispatchEvent(before));
  if (before.defaultPrevented) return;
  const sel = document.getSelection()!;
  const node = sel.anchorNode as Text;
  const at = sel.anchorOffset;
  node.data = node.data.slice(0, at) + text + node.data.slice(at);
  sel.collapse(node, at + text.length);
  fireEvent.input(field());
}

describe('the notes field', () => {
  test('shows each run with its formatting', async () => {
    await setup();
    const f = field();
    expect(f.getAttribute('contenteditable')).toBe('plaintext-only');
    expect(f.textContent).toBe(`${FIRST}\nThen show the chart.`);
    const bold = [...f.querySelectorAll('.pptx-notes__run--bold')].map((s) => s.textContent);
    expect(bold).toEqual(['headline number']);
    expect(f.querySelector('.pptx-notes__run--italic')?.textContent).toBe('pause');
    expect(f.querySelector('.pptx-notes__run--underline')?.textContent).toBe(
      'Then show the chart.'
    );
  });

  test('typed text keeps the formatting of the place it goes, before and after it is written', async () => {
    const editor = await setup();
    // Inside "headline number" (bold).
    select(18);
    typeAtCaret('XY');
    expect(field().querySelector('.pptx-notes__run--bold')?.textContent).toBe('headXYline number');
    fireEvent.blur(field());
    expect(runs(editor, 0)).toEqual([
      'Open with the ',
      'headXYline number*',
      ' and ',
      'pause',
      '.',
    ]);
    expect(field().querySelector('.pptx-notes__run--bold')?.textContent).toBe('headXYline number');
  });

  test('Ctrl+B bolds the selection as one undo step; focus, selection, and the B state stay', async () => {
    const editor = await setup();
    select(FIRST.length + 1, FIRST.length + 5);
    expect(button('Notes bold (Ctrl+B)').getAttribute('aria-pressed')).toBe('false');
    fireEvent.keyDown(field(), { key: 'b', ctrlKey: true });
    expect(runs(editor, 1)).toEqual(['Then*', ' show the chart.']);
    expect(document.activeElement).toBe(field());
    expect(notesSelection(field())).toEqual({ start: FIRST.length + 1, end: FIRST.length + 5 });
    expect(button('Notes bold (Ctrl+B)').getAttribute('aria-pressed')).toBe('true');
    // Ctrl+Z in the field undoes the format: the deck's history, not the browser's.
    fireEvent.keyDown(field(), { key: 'z', ctrlKey: true });
    expect(runs(editor, 1)).toEqual(['Then show the chart.']);
    fireEvent.keyDown(field(), { key: 'z', ctrlKey: true, shiftKey: true });
    expect(runs(editor, 1)).toEqual(['Then*', ' show the chart.']);
  });

  test('the first format on plain notes keeps the field and its focus', async () => {
    const editor = await setup(0);
    act(() => void editor.exec({ type: 'setNotes', index: 0, text: 'Plain notes here' }));
    select(0, 5);
    fireEvent.keyDown(field(), { key: 'b', metaKey: true });
    expect(document.activeElement).toBe(field());
    expect(field().querySelector('.pptx-notes__run--bold')?.textContent).toBe('Plain');
  });

  test('typed text and a format in one visit both land, as two undo steps', async () => {
    const editor = await setup();
    select(FIRST.length + 1);
    typeAtCaret('Now: ');
    select(FIRST.length + 1, FIRST.length + 4);
    fireEvent.keyDown(field(), { key: 'i', metaKey: true });
    const shown = editor.snapshot().notesFormatted![1]!.runs;
    expect(shown.map((r) => [r.text, r.italic, r.underline])).toEqual([
      ['Now', true, true],
      [': Then show the chart.', false, true],
    ]);
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().notes.endsWith('Now: Then show the chart.')).toBe(true);
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().notes.endsWith('\nThen show the chart.')).toBe(true);
  });

  test('a host button on useEditorCommand formats the notes too, and execute says if it ran', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const ran: boolean[] = [];
    function HostBold() {
      const bold = useEditorCommand('text.bold');
      return (
        <button
          type="button"
          aria-disabled={!bold.isEnabled}
          aria-pressed={bold.isActive}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => ran.push(bold.execute())}
        >
          Host bold
        </button>
      );
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <HostBold />
        <NotesPane />
      </PptxEditorContext.Provider>
    );
    const host = screen.getByRole('button', { name: 'Host bold' });
    expect(host.getAttribute('aria-disabled')).toBe('true');
    select(FIRST.length + 1, FIRST.length + 5);
    expect(host.getAttribute('aria-disabled')).toBe('false');
    fireEvent.click(host);
    expect(runs(editor, 1)).toEqual(['Then*', ' show the chart.']);
    expect(host.getAttribute('aria-pressed')).toBe('true');
    // Between two spaces there is no word: the format is armed for the next typed text.
    act(() => void editor.exec({ type: 'setNotes', index: 1, text: 'a  b' }));
    select(2);
    const was = host.getAttribute('aria-pressed');
    fireEvent.click(host);
    expect(ran).toEqual([true, true]);
    expect(host.getAttribute('aria-pressed')).toBe(was === 'true' ? 'false' : 'true');
    // Out of the notes, no shape is selected: the engine refuses Bold, and execute says so.
    fireEvent.blur(field());
    expect(host.getAttribute('aria-disabled')).toBe('true');
    fireEvent.click(host);
    expect(ran).toEqual([true, true, false]);
  });

  test('Escape goes from the notes to the current slide of the rail', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    render(
      <PptxEditorContext.Provider value={editor}>
        <div className="pptx-editor">
          <SlideRail />
          <NotesPane />
        </div>
      </PptxEditorContext.Provider>
    );
    select(0);
    fireEvent.keyDown(field(), { key: 'Escape' });
    expect(
      document.activeElement === document.querySelector('.pptx-rail__item[data-index="1"]')
    ).toBe(true);
  });

  test('Escape goes to the section header when the current slide is in a collapsed section', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const opening = editor.snapshot().sections![0]!.id;
    editor.exec({ type: 'setSectionsCollapsed', collapsed: true, sectionIds: [opening] });
    render(
      <PptxEditorContext.Provider value={editor}>
        <div className="pptx-editor">
          <SlideRail />
          <NotesPane />
        </div>
      </PptxEditorContext.Provider>
    );
    expect(document.querySelector('.pptx-rail__item[data-index="1"]') === null).toBe(true);
    select(0);
    fireEvent.keyDown(field(), { key: 'Escape' });
    const header = screen.getByRole('treeitem', { name: /^Section Opening,/ });
    expect(header.getAttribute('aria-expanded')).toBe('false');
    expect(document.activeElement === header).toBe(true);
  });

  test('Ctrl+M in the notes keeps the typed text and adds a slide; Cmd+M does not', async () => {
    const editor = await setup();
    const count = editor.snapshot().slides.length;
    select(FIRST.length);
    typeAtCaret(' Smile.');
    fireEvent.keyDown(field(), { key: 'm', metaKey: true });
    expect(editor.snapshot().slides.length).toBe(count);
    fireEvent.keyDown(field(), { key: 'm', ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(count + 1);
    editor.goToSlide(1);
    expect(editor.snapshot().notes.startsWith(`${FIRST} Smile.`)).toBe(true);
  });

  test('Enter starts a new paragraph', async () => {
    const editor = await setup();
    select(FIRST.length);
    fireEvent.keyDown(field(), { key: 'Enter' });
    fireEvent.blur(field());
    expect(editor.snapshot().notes).toBe(`${FIRST}\n\nThen show the chart.`);
  });

  test('Enter at the end adds the end marker, so the new last line shows', async () => {
    const editor = await setup();
    const all = `${FIRST}\nThen show the chart.`;
    select(all.length);
    fireEvent.keyDown(field(), { key: 'Enter' });
    // One end marker, and nothing after it in the field.
    const markers = field().querySelectorAll('br[data-end]');
    expect(markers.length).toBe(1);
    const rest = document.createRange();
    rest.setStartAfter(markers[0]!);
    rest.setEnd(field(), field().childNodes.length);
    expect(rest.toString()).toBe('');
    expect(notesFieldText(field())).toBe(`${all}\n`);
    expect(notesSelection(field())).toEqual({ start: all.length + 1, end: all.length + 1 });
    fireEvent.blur(field());
    expect(editor.snapshot().notes).toBe(`${all}\n`);
  });

  test('a character that setNotes drops does not move the formatted range', async () => {
    const editor = await setup();
    select(0);
    typeAtCaret('\u0001');
    select(1, 5);
    // A real click: the pointer goes down first, which commits the draft and redraws the field.
    fireEvent.pointerDown(button('Notes bold (Ctrl+B)'));
    expect(notesSelection(field())).toEqual({ start: 0, end: 4 });
    fireEvent.click(button('Notes bold (Ctrl+B)'));
    expect(runs(editor, 0)[0]).toBe('Open*');
    expect(editor.snapshot().notes.startsWith('Open')).toBe(true);
    expect(notesSelection(field())).toEqual({ start: 0, end: 4 });
  });

  test('a caret at the start of a paragraph reads the formats of its first character', async () => {
    await setup();
    const second = FIRST.length + 1;
    expect(notesFormatsAt(field(), { start: second, end: second }).underline).toBe(true);
    // Inside the paragraph, the character before the caret counts.
    expect(notesFormatsAt(field(), { start: second + 2, end: second + 2 }).underline).toBe(true);
    expect(notesFormatsAt(field(), { start: second - 1, end: second - 1 }).underline).toBe(false);
  });

  test('with a caret in a word, B bolds the word and the caret stays', async () => {
    const editor = await setup();
    select(2);
    fireEvent.click(button('Notes bold (Ctrl+B)'));
    expect(runs(editor, 0)[0]).toBe('Open*');
    expect(notesSelection(field())).toEqual({ start: 2, end: 2 });
    // At the end of a word the caret is not in it: no word changes, and Italic is armed.
    select(4);
    fireEvent.keyDown(field(), { key: 'i', ctrlKey: true });
    expect(screen.queryByRole('alert')).toBe(null);
    expect(button('Notes italic (Ctrl+I)').getAttribute('aria-pressed')).toBe('true');
    expect(
      editor.snapshot().notesFormatted![0]!.runs.some((r) => r.italic && r.text === 'Open')
    ).toBe(false);
  });

  test('with a caret in no word, B, I, and U stay on and arm the format', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="text.bold" />
        </Toolbar>
        <NotesPane />
      </PptxEditorContext.Provider>
    );
    const toolbarBold = (): HTMLElement =>
      container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    // At the end of a word: no word to format, but the next typed text can take one.
    select(4);
    for (const name of [
      'Notes bold (Ctrl+B)',
      'Notes italic (Ctrl+I)',
      'Notes underline (Ctrl+U)',
    ]) {
      expect(button(name).hasAttribute('aria-disabled')).toBe(false);
    }
    expect(toolbarBold().hasAttribute('aria-disabled')).toBe(false);
    // A click arms Bold: no text changes until the person types.
    fireEvent.click(button('Notes bold (Ctrl+B)'));
    expect(runs(editor, 1)).toEqual(['Then show the chart.']);
    expect(button('Notes bold (Ctrl+B)').getAttribute('aria-pressed')).toBe('true');
    // With a selection, they take a pick.
    select(0, 4);
    expect(button('Notes italic (Ctrl+I)').hasAttribute('aria-disabled')).toBe(false);
  });

  test('a blur commit puts no selection back into the field, so the focus can leave', async () => {
    await setup();
    select(FIRST.length);
    typeAtCaret(' Now.');
    const sel = document.getSelection()!;
    const addRange = sel.addRange.bind(sel);
    let added = 0;
    sel.addRange = (range: Range) => {
      added += 1;
      addRange(range);
    };
    try {
      // Tab: the focus leaves the field, and the blur writes the draft.
      act(() => field().blur());
      expect(notesFieldText(field())).toBe(`${FIRST} Now.\nThen show the chart.`);
      expect(added).toBe(0);
      expect(document.activeElement).not.toBe(field());
    } finally {
      sel.addRange = addRange;
    }
  });

  test('a selection change while focused updates B, I, and U; a blur clears them', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="text.bold" />
        </Toolbar>
        <NotesPane />
      </PptxEditorContext.Provider>
    );
    const toolbarBold = (): HTMLElement =>
      container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    const bold = (): HTMLElement => button('Notes bold (Ctrl+B)');
    // Out of the field, B, I, and U are off.
    expect(bold().getAttribute('aria-disabled')).toBe('true');
    // A caret at the end of a word: they arm the next typed text, so they are on.
    select(4);
    expect(bold().hasAttribute('aria-disabled')).toBe(false);
    // A script selects a word: no key or pointer event, only `selectionchange`.
    act(() => {
      selectNotes(field(), 0, 4);
      document.dispatchEvent(new Event('selectionchange'));
    });
    expect(bold().hasAttribute('aria-disabled')).toBe(false);
    expect(toolbarBold().hasAttribute('aria-disabled')).toBe(false);
    act(() => field().blur());
    expect(bold().getAttribute('aria-disabled')).toBe('true');
    expect(bold().getAttribute('aria-pressed')).toBe('false');
  });

  test('focus that comes back with the same selection enables B, I, and U again', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="text.bold" />
        </Toolbar>
        <NotesPane />
      </PptxEditorContext.Provider>
    );
    const toolbarBold = (): HTMLElement =>
      container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    const bold = (): HTMLElement => button('Notes bold (Ctrl+B)');
    select(0, 4);
    expect(bold().hasAttribute('aria-disabled')).toBe(false);
    act(() => field().blur());
    expect(bold().getAttribute('aria-disabled')).toBe('true');
    // The window or a script gives the focus back: the selection is the same, so the browser
    // fires no selectionchange, and no key or pointer event comes.
    act(() => {
      selectNotes(field(), 0, 4);
      field().focus();
    });
    expect(bold().hasAttribute('aria-disabled')).toBe(false);
    expect(toolbarBold().hasAttribute('aria-disabled')).toBe(false);
  });

  test('focus in the notes leaves no shape selected, and toolbar Bold formats the notes', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar preset={false}>
          <ToolbarSlot slotId="text.bold" />
        </Toolbar>
        <NotesPane />
      </PptxEditorContext.Provider>
    );
    const title = editor.session!.layout(1)!.items[0]!.element.nodeId;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [title] }));
    select(FIRST.length + 1, FIRST.length + 5);
    expect(editor.snapshot().selection).toEqual({ kind: 'none' });
    const bold = container.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(bold.hasAttribute('data-disabled')).toBe(false);
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    fireEvent.click(bold);
    expect(runs(editor, 1)).toEqual(['Then*', ' show the chart.']);
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    // Undo from the field keeps no shape selected behind the notes.
    fireEvent.keyDown(field(), { key: 'z', ctrlKey: true });
    expect(runs(editor, 1)).toEqual(['Then show the chart.']);
    expect(editor.snapshot().selection).toEqual({ kind: 'none' });
    // Out of the notes, Bold is the slide's again.
    fireEvent.blur(field());
    expect(bold.hasAttribute('data-disabled')).toBe(true);
  });

  test('an emptied field whose browser keeps a placeholder <br> saves no notes', async () => {
    const editor = await setup();
    act(() => field().focus());
    field().replaceChildren(document.createElement('br'));
    fireEvent.input(field());
    fireEvent.blur(field());
    expect(editor.snapshot().notes).toBe('');
  });

  test('showFormatButtons={false} hides the buttons; the keys still format', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <NotesPane className="acme-notes" showFormatButtons={false} />
      </PptxEditorContext.Provider>
    );
    expect(container.querySelector('.pptx-notes.acme-notes') !== null).toBe(true);
    expect(screen.queryByRole('button', { name: 'Notes bold (Ctrl+B)' }) === null).toBe(true);
    select(0, 4);
    fireEvent.keyDown(field(), { key: 'b', ctrlKey: true });
    expect(runs(editor, 0)[0]).toBe('Open*');
  });

  test('a selection and the B button bold it', async () => {
    const editor = await setup();
    select(0, 4);
    fireEvent.click(button('Notes bold (Ctrl+B)'));
    expect(runs(editor, 0)[0]).toBe('Open*');
  });

  test('viewing mode shows the notes read-only, without format buttons', async () => {
    const editor = await setup();
    act(() => editor.setMode('view'));
    expect(field().getAttribute('aria-readonly')).toBe('true');
    expect(field().getAttribute('contenteditable')).toBe('false');
    expect(screen.queryByRole('button', { name: 'Notes bold (Ctrl+B)' }) === null).toBe(true);
    select(0, 4);
    fireEvent.keyDown(field(), { key: 'b', ctrlKey: true });
    expect(runs(editor, 0)[0]).toBe('Open with the ');
  });
});

describe('the preset toolbar while the notes field has focus', () => {
  test('shows Undo, Redo, the notes formats, and Find; strikethrough formats the notes', async () => {
    cleanup();
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    editor.goToSlide(1);
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
        <NotesPane />
      </PptxEditorContext.Provider>
    );
    select(0, 4);
    const slots = [...container.querySelectorAll('.pptx-toolbar [data-slot]')].map((e) =>
      e.getAttribute('data-slot')
    );
    expect(slots).toEqual([
      'history.undo',
      'history.redo',
      'text.bold',
      'text.italic',
      'text.underline',
      'text.strike',
      'find.find',
    ]);
    fireEvent.click(container.querySelector<HTMLElement>('[data-slot="text.strike"]')!);
    expect(
      editor.snapshot().notesFormatted![0]!.runs.some((r) => r.strike && r.text === 'Open')
    ).toBe(true);
  });
});
