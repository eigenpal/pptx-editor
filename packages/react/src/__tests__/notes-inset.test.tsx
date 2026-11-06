import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { NOTES_INSET_PROPERTY, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor } from '../index.ts';

// The default editor keeps its status area and font notice above the notes card: the views box
// carries the card's height in `--pptx-notes-inset`, in the Normal view only.

const saved = HTMLElement.prototype.getBoundingClientRect;
// happy-dom defines the method on Element.prototype. Restore by removing the own property, so an
// assignment here never shadows a later test's patch of Element.prototype in the same process.
const ownBefore = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'getBoundingClientRect');
function restoreRect(): void {
  if (ownBefore) Object.defineProperty(HTMLElement.prototype, 'getBoundingClientRect', ownBefore);
  else delete (HTMLElement.prototype as { getBoundingClientRect?: unknown }).getBoundingClientRect;
}
afterEach(() => {
  cleanup();
  restoreRect();
});

/** happy-dom lays out nothing: give the views box and the notes card their boxes. */
function layOut(): void {
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    const box = (top: number, bottom: number) =>
      ({ top, bottom, left: 0, right: 900, width: 900, height: bottom - top }) as DOMRect;
    if (this.classList.contains('pptx-default__views')) return box(0, 800);
    if (this.classList.contains('pptx-notes')) return box(650, 788);
    return saved.call(this);
  };
}

describe('the default editor and the notes card', () => {
  test('the corner parts float above the notes in the Normal view, at the corner in the sorter', async () => {
    layOut();
    let editor: PptxEditorInstance | null = null;
    render(<PptxEditor onReady={(e) => (editor = e)} />);
    await act(async () => void (await editor!.load(buildPptx({ slides: [{ shapes: '' }] }))));
    const views = document.querySelector<HTMLElement>('.pptx-default__views')!;
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('150px');
    const sorter = screen.getByRole('radio', { name: 'Slide sorter' });
    act(() => void fireEvent.click(sorter));
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
    // The status area stays mounted, so the view switch keeps its focus.
    expect(screen.getByRole('radio', { name: 'Slide sorter' })).toBe(sorter);
    act(() => void fireEvent.click(screen.getByRole('radio', { name: 'Normal' })));
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('150px');
  });

  test('without the notes pane the parts stay at the corner', async () => {
    layOut();
    let editor: PptxEditorInstance | null = null;
    render(<PptxEditor notesPane={false} onReady={(e) => (editor = e)} />);
    await act(async () => void (await editor!.load(buildPptx({ slides: [{ shapes: '' }] }))));
    const views = document.querySelector<HTMLElement>('.pptx-default__views')!;
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
  });
});
