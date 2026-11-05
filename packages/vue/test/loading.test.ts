// The loading screen: the Vue twin of React's `PptxEditor.Loading`, with the same markup and the
// same rule (`showsLoading` in core).

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, h, nextTick, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { PptxEditor, PptxEditorLoading, PptxEditorLoadingSpinner } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { shallowRef } from 'vue';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

const deck = () => buildPptx({ slides: [{ shapes: '' }] });
const screenEl = () => host!.querySelector<HTMLElement>('.pptx-loading');

function mount(editor: PptxEditorInstance, render: () => unknown): void {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp({
    setup() {
      providePptxEditor(shallowRef(editor));
      return render;
    },
  });
  app.mount(host);
}

describe('PptxEditorLoading', () => {
  test('shows a 16:9 sheet with a spinner before the deck, and nothing once it is ready', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    mount(editor, () => h(PptxEditorLoading, { overlay: true }));
    const el = screenEl()!;
    expect(el.getAttribute('role')).toBe('status');
    expect(el.classList.contains('pptx-editor')).toBe(true);
    expect(el.classList.contains('pptx-loading--overlay')).toBe(true);
    expect(el.querySelector('.pptx-loading__spinner') !== null).toBe(true);
    expect(el.textContent).toBe('Opening presentation…');
    expect(el.querySelector<HTMLElement>('.pptx-loading__sheet')!.style.aspectRatio).toBe(
      '1280 / 720'
    );
    expect((await editor.load(deck())).ok).toBe(true);
    await nextTick();
    expect(screenEl() === null).toBe(true);
    // The next deck: the screen covers the open one until it is ready.
    const opening = editor.load(deck());
    await nextTick();
    expect(screenEl() !== null).toBe(true);
    await opening;
    await nextTick();
    expect(screenEl() === null).toBe(true);
  });

  test('a refused file clears the screen, so it never spins forever', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    mount(editor, () => h(PptxEditorLoading));
    await editor.load(new Uint8Array([1, 2, 3]));
    await nextTick();
    expect(editor.snapshot().status).toBe('error');
    expect(screenEl() === null).toBe(true);
  });

  test('the indicator slot replaces the spinner; the default slot replaces the sheet', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    mount(editor, () => [
      h(PptxEditorLoading, null, { indicator: () => h('b', { id: 'mine' }, 'Wait') }),
      h(PptxEditorLoading, { class: 'custom', when: true }, () => h(PptxEditorLoadingSpinner)),
    ]);
    expect(host!.querySelector('.pptx-loading__sheet #mine') !== null).toBe(true);
    expect(host!.querySelector('.custom .pptx-loading__spinner') !== null).toBe(true);
    await editor.load(deck());
    await nextTick();
    expect(host!.querySelector('#mine') === null).toBe(true);
    expect(host!.querySelector('.custom') !== null).toBe(true);
  });

  test('the default editor shows the loading screen until the deck is ready', async () => {
    let editor: PptxEditorInstance | null = null;
    host = document.createElement('div');
    document.body.append(host);
    app = createApp({
      render: () => h(PptxEditor, { onReady: (e: PptxEditorInstance) => (editor = e) }),
    });
    app.mount(host);
    await nextTick();
    expect(host.querySelector('.pptx-default__main .pptx-loading--overlay') !== null).toBe(true);
    await editor!.load(deck());
    await nextTick();
    expect(screenEl() === null).toBe(true);
  });
});
