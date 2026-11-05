import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor } from '../index.ts';

// The loading screen: shown while the editor has no deck and while a deck opens, cleared when
// the deck is ready or the file is refused.

afterEach(() => cleanup());

const deck = () => buildPptx({ slides: [{ shapes: '' }] });
const screenEl = () => document.querySelector<HTMLElement>('.pptx-loading');

function mountRoot(ui: React.ReactNode = <PptxEditor.Loading overlay />) {
  let editor: PptxEditorInstance | null = null;
  render(<PptxEditor.Root onReady={(e) => (editor = e)}>{ui}</PptxEditor.Root>);
  return () => editor!;
}

describe('PptxEditor.Loading', () => {
  test('shows a 16:9 sheet with a spinner before the deck, and nothing once it is ready', async () => {
    const editor = mountRoot();
    const el = screenEl()!;
    expect(el.getAttribute('role')).toBe('status');
    expect(el.classList.contains('pptx-editor')).toBe(true);
    expect(el.classList.contains('pptx-loading--overlay')).toBe(true);
    expect(el.querySelector('.pptx-loading__spinner') !== null).toBe(true);
    expect(el.textContent).toBe('Opening presentation…');
    const sheet = el.querySelector<HTMLElement>('.pptx-loading__sheet')!;
    expect(sheet.style.aspectRatio).toBe('1280 / 720');
    await act(async () => void (await editor().load(deck())));
    expect(screenEl() === null).toBe(true);
  });

  test('covers the open deck while the next one opens', async () => {
    const editor = mountRoot();
    await act(async () => void (await editor().load(deck())));
    let opening!: Promise<unknown>;
    act(() => void (opening = editor().load(deck())));
    expect(editor().snapshot().status).toBe('loading');
    expect(screenEl() !== null).toBe(true);
    await act(async () => void (await opening));
    expect(screenEl() === null).toBe(true);
  });

  test('a refused file clears the screen, so it never spins forever', async () => {
    const editor = mountRoot();
    await act(async () => void (await editor().load(new Uint8Array([1, 2, 3]))));
    expect(editor().snapshot().status).toBe('error');
    expect(screenEl() === null).toBe(true);
  });

  test('indicator replaces the spinner, children replace the sheet, and when holds it', async () => {
    const editor = mountRoot(
      <>
        <PptxEditor.Loading indicator={<b id="mine">Wait</b>} />
        <PptxEditor.Loading className="custom" when>
          <PptxEditor.Loading.Spinner className="host-spinner" />
        </PptxEditor.Loading>
      </>
    );
    expect(document.querySelector('.pptx-loading__sheet #mine') !== null).toBe(true);
    expect(document.querySelector('.custom .host-spinner.pptx-loading__spinner') !== null).toBe(
      true
    );
    await act(async () => void (await editor().load(deck())));
    expect(document.querySelector('#mine') === null).toBe(true);
    expect(document.querySelector('.custom') !== null).toBe(true);
  });

  test('the default editor shows the loading screen until the deck is ready', async () => {
    let editor: PptxEditorInstance | null = null;
    render(<PptxEditor onReady={(e) => (editor = e)} />);
    expect(document.querySelector('.pptx-default__main .pptx-loading--overlay') !== null).toBe(
      true
    );
    await act(async () => void (await editor!.load(deck())));
    expect(screenEl() === null).toBe(true);
  });
});
