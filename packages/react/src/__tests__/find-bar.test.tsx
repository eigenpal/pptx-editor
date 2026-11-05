import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { FindBar, requestFindBar } from '../FindBar.tsx';

// The find bar keys: Ctrl/Cmd+F opens find, Ctrl/Cmd+H opens replace, Enter and Shift+Enter
// step, Escape closes and gives the focus back, and the counter reads "n of m".

const IN = 914_400;

async function setup(): Promise<{ editor: PptxEditorInstance; before: HTMLButtonElement }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const text = para(run('one two one')) + para(run('three one'));
  const r = await editor.load(
    buildPptx({
      slides: [
        { shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: IN, text }) },
        {
          shapes: shape({
            id: 2,
            name: 'U',
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            text: para(run('one')),
          }),
        },
      ],
    })
  );
  expect(r.ok).toBe(true);
  render(
    <PptxEditorContext.Provider value={editor}>
      <div className="pptx-editor">
        <button type="button">before</button>
        <FindBar />
      </div>
    </PptxEditorContext.Provider>
  );
  const before = screen.getByRole('button', { name: 'before' }) as HTMLButtonElement;
  before.focus();
  return { editor, before };
}

const input = (name: string): HTMLInputElement =>
  screen.getByRole('textbox', { name }) as HTMLInputElement;

afterEach(cleanup);

describe('FindBar', () => {
  test('Ctrl+F opens find with the focus in the field; Escape closes and returns the focus', async () => {
    const { before } = await setup();
    expect(screen.queryByRole('textbox', { name: 'Find' }) === null).toBe(true);
    fireEvent.keyDown(before, { key: 'f', ctrlKey: true });
    expect(document.activeElement).toBe(input('Find'));
    expect(screen.queryByRole('textbox', { name: 'Replace' }) === null).toBe(true);
    fireEvent.keyDown(input('Find'), { key: 'Escape' });
    expect(screen.queryByRole('textbox', { name: 'Find' }) === null).toBe(true);
    expect(document.activeElement).toBe(before);
  });

  test('Ctrl/Cmd+F after work on the host page stays with the browser', async () => {
    const { before } = await setup();
    const outside = document.createElement('input');
    document.body.appendChild(outside);
    const ctrlF = (target: EventTarget): boolean => {
      const key = new KeyboardEvent('keydown', {
        key: 'f',
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      target.dispatchEvent(key);
      return key.defaultPrevented;
    };
    try {
      // The focus on a host field, then on nothing: the browser keeps the key both times.
      outside.focus();
      expect(ctrlF(outside)).toBe(false);
      outside.blur();
      expect(ctrlF(document.body)).toBe(false);
      expect(screen.queryByRole('textbox', { name: 'Find' }) === null).toBe(true);
      // Work in the editor, then the focus drops to nothing (a closed menu): the bar opens.
      before.focus();
      before.blur();
      let opened = false;
      act(() => void (opened = ctrlF(document.body)));
      expect(opened).toBe(true);
      expect(document.activeElement).toBe(input('Find'));
    } finally {
      outside.remove();
    }
  });

  test('after the bar opens and closes, Ctrl+F with the focus on nothing opens it again', async () => {
    const { before } = await setup();
    fireEvent.keyDown(before, { key: 'f', ctrlKey: true });
    fireEvent.keyDown(input('Find'), { key: 'Escape' });
    expect(document.activeElement).toBe(before);
    // A closed menu or a removed control drops the focus to nothing, with no new press.
    before.blur();
    const key = new KeyboardEvent('keydown', {
      key: 'f',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    });
    act(() => void document.body.dispatchEvent(key));
    expect(key.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(input('Find'));
  });

  test('the bar keeps its height free above the slide while it is mounted', async () => {
    await setup();
    const scope = document.querySelector<HTMLElement>('.pptx-find')!.parentElement!;
    // Hidden, the bar reserves nothing; core's `reserveFindBarSpace` follows its height.
    expect(scope.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
    cleanup();
    expect(scope.style.getPropertyValue('--pptx-find-reserve')).toBe('');
  });

  test('Enter and Shift+Enter step through matches, with an "n of m" counter', async () => {
    const { editor, before } = await setup();
    fireEvent.keyDown(before, { key: 'f', metaKey: true });
    fireEvent.change(input('Find'), { target: { value: 'one' } });
    // Find as you type: the first match is selected at once.
    expect(screen.getByText('1 of 4')).toBeDefined();
    fireEvent.keyDown(input('Find'), { key: 'Enter' });
    expect(screen.getByText('2 of 4')).toBeDefined();
    const sel = editor.snapshot().selection;
    expect(sel.kind === 'text' && [sel.anchor, sel.focus]).toEqual([
      { paragraph: 0, offset: 8 },
      { paragraph: 0, offset: 11 },
    ]);
    fireEvent.keyDown(input('Find'), { key: 'Enter', shiftKey: true });
    expect(screen.getByText('1 of 4')).toBeDefined();
    fireEvent.keyDown(input('Find'), { key: 'Enter', shiftKey: true });
    expect(screen.getByText('4 of 4')).toBeDefined();
    expect(editor.snapshot().currentSlide).toBe(1);
    fireEvent.change(input('Find'), { target: { value: 'zebra' } });
    expect(screen.getByText('No results')).toBeDefined();
  });

  test('Ctrl+H opens replace; Enter replaces the shown match and Replace all the rest', async () => {
    const { editor, before } = await setup();
    fireEvent.keyDown(before, { key: 'h', ctrlKey: true });
    expect(document.activeElement).toBe(input('Find'));
    fireEvent.change(input('Find'), { target: { value: 'one' } });
    fireEvent.change(input('Replace'), { target: { value: 'uno' } });
    fireEvent.keyDown(input('Replace'), { key: 'Enter' });
    expect(editor.find.matches('uno')).toHaveLength(1);
    expect(screen.getByText('1 of 3')).toBeDefined();
    fireEvent.click(screen.getByRole('button', { name: 'Replace all' }));
    expect(editor.find.matches('uno')).toHaveLength(4);
    expect(screen.getByText('3 replaced')).toBeDefined();
    // Replace all is one undo step.
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.find.matches('uno')).toHaveLength(1);
  });

  test('typing selects the first match from the caret, and a longer query keeps it', async () => {
    const { editor, before } = await setup();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    // The caret after "one t": the first match from there is "one" at offset 8.
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 5 },
        focus: { paragraph: 0, offset: 5 },
      })
    );
    fireEvent.keyDown(before, { key: 'f', ctrlKey: true });
    fireEvent.change(input('Find'), { target: { value: 'o' } });
    expect(screen.getByText('2 of 5')).toBeDefined();
    fireEvent.change(input('Find'), { target: { value: 'on' } });
    expect(screen.getByText('2 of 4')).toBeDefined();
    fireEvent.change(input('Find'), { target: { value: 'one' } });
    expect(screen.getByText('2 of 4')).toBeDefined();
    const sel = editor.snapshot().selection;
    expect(sel.kind === 'text' && [sel.anchor, sel.focus]).toEqual([
      { paragraph: 0, offset: 8 },
      { paragraph: 0, offset: 11 },
    ]);
    expect(document.activeElement).toBe(input('Find'));
  });

  test('a toolbar or menu request opens the bar of its own editor only', async () => {
    const { editor } = await setup();
    const other = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    act(() => requestFindBar(other));
    expect(screen.queryByRole('textbox', { name: 'Find' }) === null).toBe(true);
    act(() => requestFindBar(editor, true));
    expect(screen.getByRole('textbox', { name: 'Replace' })).toBeDefined();
  });
});
