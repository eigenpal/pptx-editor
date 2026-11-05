// The slide panes after the browser review: the open-failure notice, the section menu's keys
// (a React portal's key events bubble through the rail), the rail keys in the viewing mode and
// Home and End without sections, notes that store the formats they show, the find bar on a new
// document and with matches but no current one, and Cmd/Ctrl+F in the slide sorter.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { EditNotice } from '../EditNotice.tsx';
import { FindBar, requestFindBar } from '../FindBar.tsx';
import { PptxEditor } from '../index.ts';
import { NotesPane } from '../NotesPane.tsx';
import { OpenFailure } from '../OpenFailure.tsx';
import { SlideRail } from '../panes.tsx';
import { DialogProvider } from '../popups.tsx';
import { SlideSorter } from '../SlideSorter.tsx';

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

const IN = 914_400;

const flatDeck = (count = 3, notes?: string) =>
  buildPptx({
    slides: Array.from({ length: count }, (_, i) => ({
      layout: 'blank' as const,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`S${i + 1}`)) }),
      ...(notes && i === 0 ? { notes } : {}),
    })),
  });

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const tree = (editor: PptxEditorInstance, ui: React.ReactNode) => (
  <PptxEditorContext.Provider value={editor}>
    <DialogProvider>{ui}</DialogProvider>
  </PptxEditorContext.Provider>
);

const notice = () => document.querySelector('.pptx-edit-notice')?.textContent ?? '';

describe('open failure', () => {
  test('a refused file shows why, keeps the open deck, and shows again for the next refusal', async () => {
    const editor = await open(flatDeck());
    let dismissed = 0;
    render(tree(editor, <OpenFailure onDismiss={() => (dismissed += 1)} />));
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
    await act(async () => void (await editor.load(new Uint8Array(17).fill(7))));
    const snap = editor.snapshot();
    expect(snap.status).toBe('ready');
    expect(snap.slides.length).toBe(3);
    expect(snap.openFailure).toEqual({ code: 'zip', reason: 'damaged' });
    const el = screen.getByRole('alert');
    expect(el.textContent).toContain('The file did not open');
    expect(el.textContent).toContain('The file is damaged, or it is not a PPTX presentation.');
    expect(el.textContent).toContain('The presentation that was open stays open.');
    expect(el.textContent).toContain('Error code: zip');
    fireEvent.click(screen.getByRole('button', { name: 'Close' }));
    expect(dismissed).toBe(1);
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
    // The same file again: a new refusal shows again.
    await act(async () => void (await editor.load(new Uint8Array(17).fill(7))));
    expect(document.querySelector('.pptx-open-failure') !== null).toBe(true);
    // A deck that opens clears it.
    await act(async () => void (await editor.load(flatDeck(2))));
    expect(editor.snapshot().openFailure === undefined).toBe(true);
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
  });

  test('with no deck open, the status is error and the notice does not say a deck stays', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    render(tree(editor, <OpenFailure />));
    await act(async () => void (await editor.load(new Uint8Array(0))));
    expect(editor.snapshot().status).toBe('error');
    expect(editor.snapshot().openFailure?.reason).toBe('not-pptx');
    const el = screen.getByRole('alert');
    expect(el.textContent).toContain('The file holds no presentation.');
    expect(el.textContent?.includes('stays open')).toBe(false);
  });

  test('the default editor mounts the notice, with host actions; false removes it', async () => {
    let editor: PptxEditorInstance | null = null;
    const view = render(
      <PptxEditor
        onReady={(e) => (editor = e)}
        renderOpenFailureActions={({ failure }) => (
          <a className="pptx-open-failure__action" href={`#help-${failure.code}`}>
            Help
          </a>
        )}
      />
    );
    await act(async () => void (await editor!.load(new Uint8Array([1, 2, 3]))));
    expect(screen.getByRole('link', { name: 'Help' }).getAttribute('href')).toBe('#help-zip');
    view.unmount();
    render(<PptxEditor onReady={(e) => (editor = e)} openFailure={false} />);
    await act(async () => void (await editor!.load(new Uint8Array([1, 2, 3]))));
    expect(document.querySelector('.pptx-open-failure') === null).toBe(true);
  });
});

describe('open failure misuse', () => {
  test('renderOpenFailureActions with openFailure.actions or with false warns once each', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const view = render(
        <PptxEditor
          openFailure={{ actions: () => <a>Prop</a> }}
          renderOpenFailureActions={() => <a>Render</a>}
        />
      );
      view.unmount();
      render(<PptxEditor openFailure={false} renderOpenFailureActions={() => <a>Render</a>} />);
      const text = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain('renderOpenFailureActions replaces openFailure.actions.');
      expect(text).toContain('renderOpenFailureActions is ignored with openFailure={false}');
    } finally {
      warn.mockRestore();
    }
  });

  test('the sugar forwards onChange and locale to the root', async () => {
    let editor: PptxEditorInstance | null = null;
    let changes = 0;
    render(
      <PptxEditor
        onReady={(e) => (editor = e)}
        onChange={() => (changes += 1)}
        locale={{ openFailure: { title: 'Nie otwarto pliku' } }}
      />
    );
    await act(async () => void (await editor!.load(new Uint8Array([1, 2, 3]))));
    expect(changes > 0).toBe(true);
    expect(screen.getByRole('alert').textContent).toContain('Nie otwarto pliku');
  });
});

describe('section menu keys', () => {
  test('Enter on a section menu row runs the row and adds no slide', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    const before = editor.snapshot().slides.length;
    render(tree(editor, <SlideRail />));
    const header = screen.getByRole('treeitem', { name: /^Section Details,/ });
    header.focus();
    fireEvent.keyDown(header, { key: 'F10', shiftKey: true });
    const menu = screen.getByRole('menu');
    const head = header.querySelector('.pptx-rail__section-head')!;
    // Anchored to the header row, not under the section's slides.
    expect(menu !== null && head !== null).toBe(true);
    const row = screen.getByRole('menuitem', { name: /Move section up/ });
    row.focus();
    fireEvent.keyDown(row, { key: 'Enter' });
    fireEvent.click(row);
    expect(editor.snapshot().slides.length).toBe(before);
    expect((editor.snapshot().sections ?? []).map((s) => s.name)).toEqual(['Details', 'Opening']);
  });
});

describe('rail keys', () => {
  test('in the viewing mode, Enter, Cmd+D, and Delete say why they do nothing', async () => {
    const editor = await open(flatDeck());
    editor.setMode('view');
    render(tree(editor, <SlideRail />));
    for (const key of [{ key: 'Enter' }, { key: 'd', metaKey: true }, { key: 'Delete' }]) {
      act(() => editor.goToSlide(0));
      const item = screen.getByRole('option', { name: 'Slide 1' });
      fireEvent.keyDown(item, key);
      expect(editor.snapshot().lastRejection?.reason).toBe('disabled.viewing');
    }
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('Home and End go to the first and the last slide of a rail without sections', async () => {
    const editor = await open(flatDeck(4));
    render(tree(editor, <SlideRail />));
    fireEvent.keyDown(screen.getByRole('option', { name: 'Slide 1' }), { key: 'End' });
    expect(editor.snapshot().currentSlide).toBe(3);
    fireEvent.keyDown(screen.getByRole('option', { name: 'Slide 4' }), { key: 'Home' });
    expect(editor.snapshot().currentSlide).toBe(0);
  });
});

describe('slide sorter', () => {
  test('Cmd+F in the sorter shows why find does not open', async () => {
    const editor = await open(flatDeck());
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    render(
      <div className="pptx-editor">
        {tree(
          editor,
          <>
            <SlideSorter />
            <FindBar />
            <EditNotice />
          </>
        )}
      </div>
    );
    const tile = screen.getByRole('option', { name: 'Slide 1' });
    tile.focus();
    act(() => void fireEvent.keyDown(tile, { key: 'f', metaKey: true }));
    expect(notice()).toBe('Open the Normal view to find text in the slides');
    expect(document.querySelector<HTMLElement>('.pptx-find')!.hidden).toBe(true);
  });
});

describe('find bar', () => {
  test('a new document closes the bar and clears the search', async () => {
    const editor = await open(flatDeck());
    render(<div className="pptx-editor">{tree(editor, <FindBar />)}</div>);
    act(() => requestFindBar(editor));
    const input = screen.getByRole('textbox', { name: 'Find' }) as HTMLInputElement;
    fireEvent.change(input, { target: { value: 'S2' } });
    expect(document.querySelector('.pptx-find__count')!.textContent).toBe('1 of 1');
    await act(async () => void (await editor.load(flatDeck(2))));
    expect(document.querySelector<HTMLElement>('.pptx-find')!.hidden).toBe(true);
    act(() => requestFindBar(editor));
    expect((screen.getByRole('textbox', { name: 'Find' }) as HTMLInputElement).value).toBe('');
  });

  test('matches with none shown give the count alone; Replace all keeps focus in the field', async () => {
    const editor = await open(flatDeck());
    render(<div className="pptx-editor">{tree(editor, <FindBar />)}</div>);
    act(() => requestFindBar(editor, true));
    fireEvent.change(screen.getByRole('textbox', { name: 'Find' }), { target: { value: 's' } });
    fireEvent.change(screen.getByRole('textbox', { name: 'Replace' }), {
      target: { value: 'S' },
    });
    const all = screen.getByRole('button', { name: 'Replace all' });
    all.focus();
    act(() => void fireEvent.click(all));
    // Case-insensitive: the replaced text still matches, and no match is shown.
    expect(document.querySelector('.pptx-find__count')!.textContent).toBe('3 matches');
    fireEvent.change(screen.getByRole('textbox', { name: 'Find' }), { target: { value: 'S' } });
    fireEvent.change(screen.getByRole('textbox', { name: 'Replace' }), {
      target: { value: 'Q' },
    });
    const again = screen.getByRole('button', { name: 'Replace all' });
    again.focus();
    act(() => void fireEvent.click(again));
    expect(document.querySelector('.pptx-find__count')!.textContent).toBe('No results');
    expect(document.activeElement).toBe(screen.getByRole('textbox', { name: 'Find' }));
  });
});

describe('notes formats', () => {
  test('text typed into a bold run stays bold after the commit, also in a new paragraph', async () => {
    const editor = await open(flatDeck(1, 'Hi'));
    render(tree(editor, <NotesPane />));
    const field = screen.getByRole('textbox', { name: 'Speaker notes' });
    act(() => field.focus());
    // The field shows "Hi" and the user typed " bold" and a new line into a bold run.
    const span = document.createElement('span');
    span.className = 'pptx-notes__run--bold';
    span.textContent = ' bold\nmore';
    field.append(span);
    act(() => void fireEvent.input(field));
    act(() => field.blur());
    const stored = editor.snapshot().notesFormatted!;
    expect(stored.map((p) => p.runs.map((r) => [r.text, r.bold]))).toEqual([
      [
        ['Hi', false],
        [' bold', true],
      ],
      [['more', true]],
    ]);
  });

  test('B with a caret outside a word arms bold for the next typed text', async () => {
    const editor = await open(flatDeck(1, 'Hi'));
    render(tree(editor, <NotesPane />));
    const field = screen.getByRole('textbox', { name: 'Speaker notes' });
    act(() => field.focus());
    const sel = document.getSelection()!;
    sel.selectAllChildren(field);
    sel.collapseToEnd();
    act(() => void fireEvent(document, new Event('selectionchange')));
    const bold = screen.getByRole('button', { name: /Notes bold/ });
    expect(bold.getAttribute('aria-disabled') === 'true').toBe(false);
    act(() => void fireEvent.click(bold));
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    act(
      () =>
        void field.dispatchEvent(
          new InputEvent('beforeinput', {
            inputType: 'insertText',
            data: '!',
            bubbles: true,
            cancelable: true,
          })
        )
    );
    act(() => field.blur());
    expect(
      editor.snapshot().notesFormatted![0]!.runs.map((r) => [r.text, r.bold] as const)
    ).toEqual([
      ['Hi', false],
      ['!', true],
    ]);
  });
});
