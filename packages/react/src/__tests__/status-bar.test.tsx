import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { resetDevWarnings, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor } from '../index.ts';

// The status bar: slide x of y, the Notes button, the view switch, and the zoom pill.

afterEach(() => cleanup());

const deck = () => buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] });

async function mount(ui: (ready: (e: PptxEditorInstance) => void) => React.ReactNode) {
  let editor: PptxEditorInstance | null = null;
  render(<>{ui((e) => (editor = e))}</>);
  await act(async () => void (await editor!.load(deck())));
  return editor!;
}

describe('PptxEditor.StatusBar', () => {
  test('slide x of y follows the current slide; the zoom pill steps and fits', async () => {
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.StatusBar />
      </PptxEditor.Root>
    ));
    expect(document.querySelector('[data-part="slideNumber"]')!.textContent).toBe('Slide 1 of 3');
    act(() => editor.selectSlide(2));
    expect(document.querySelector('[data-part="slideNumber"]')!.textContent).toBe('Slide 3 of 3');
    act(() => editor.setZoom(1));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Zoom in' })));
    expect(editor.snapshot().zoom).toBe(1.1);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Zoom: 110%' })));
    act(() => void fireEvent.click(screen.getByRole('menuitemradio', { name: 'Fit to window' })));
    expect(editor.snapshot().fitZoom).toBe(true);
  });

  test('the view switch is the editor command: Slide Sorter changes the view', async () => {
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.StatusBar />
      </PptxEditor.Root>
    ));
    const sorter = screen.getByRole('radio', { name: 'Slide sorter' });
    act(() => void fireEvent.click(sorter));
    expect(editor.snapshot().presentationView).toBe('slideSorter');
    expect(sorter.getAttribute('aria-checked')).toBe('true');
  });

  test('the Notes button and the notes pane share their open state', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.NotesPane />
        <PptxEditor.StatusBar />
      </PptxEditor.Root>
    ));
    const notes = screen.getByRole('button', { name: 'Notes' });
    expect(notes.getAttribute('aria-pressed')).toBe('true');
    expect(document.querySelector('.pptx-notes__text') !== null).toBe(true);
    act(() => void fireEvent.click(notes));
    expect(notes.getAttribute('aria-pressed')).toBe('false');
    expect(document.querySelector('.pptx-notes__text') === null).toBe(true);
  });

  test('named parts change their part in place, and preset={false} keeps only children', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.StatusBar className="acme-status">
          <PptxEditor.StatusBar.Views hidden />
          <PptxEditor.StatusBar.SlideNumber className="acme-slide" />
          <span data-testid="saved">Saved</span>
        </PptxEditor.StatusBar>
        <PptxEditor.StatusBar preset={false}>
          <PptxEditor.StatusBar.Zoom />
        </PptxEditor.StatusBar>
      </PptxEditor.Root>
    ));
    const [first, second] = [...document.querySelectorAll('.pptx-statusbar')];
    expect(first!.classList.contains('acme-status')).toBe(true);
    expect(first!.querySelector('[data-part="views"]') === null).toBe(true);
    expect(first!.querySelector('.acme-slide') !== null).toBe(true);
    expect(screen.getByTestId('saved')).toBeTruthy();
    expect(second!.querySelector('[data-part="slideNumber"]') === null).toBe(true);
    expect(second!.querySelector('[data-part="zoom"]') !== null).toBe(true);
  });

  test('Notes takes the ladder: an icon, and asChild with a host element', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.NotesPane />
        <PptxEditor.StatusBar>
          <PptxEditor.StatusBar.Notes icon={<i data-testid="host-icon" />} />
        </PptxEditor.StatusBar>
        <PptxEditor.StatusBar preset={false}>
          <PptxEditor.StatusBar.Notes asChild>
            <a href="#notes" data-testid="host-notes">
              Speaker notes
            </a>
          </PptxEditor.StatusBar.Notes>
        </PptxEditor.StatusBar>
      </PptxEditor.Root>
    ));
    expect(screen.getByTestId('host-icon')).toBeTruthy();
    const host = screen.getByTestId('host-notes');
    expect(host.getAttribute('aria-label')).toBe('Notes');
    expect(host.getAttribute('aria-pressed')).toBe('true');
    act(() => void fireEvent.click(host));
    expect(host.getAttribute('aria-pressed')).toBe('false');
    expect(document.querySelector('.pptx-notes__text') === null).toBe(true);
  });

  test('Notes with asChild and no element child warns by its part and keeps its icon', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await mount((onReady) => (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.StatusBar preset={false}>
            <PptxEditor.StatusBar.Notes asChild>text only</PptxEditor.StatusBar.Notes>
          </PptxEditor.StatusBar>
        </PptxEditor.Root>
      ));
      const button = screen.getByRole('button', { name: 'Notes' });
      expect(button.querySelector('.pptx-icon') !== null).toBe(true);
      expect(String(warn.mock.calls[0]?.[0])).toContain('StatusBar.Notes "notes" has asChild');
    } finally {
      warn.mockRestore();
    }
  });
});
