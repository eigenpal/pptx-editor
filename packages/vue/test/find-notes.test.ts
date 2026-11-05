// Find in the speaker notes, in both adapters: each Next step inside the notes moves the
// counter and paints the match in the notes field, though the slide and the selection stay.

import { afterEach, describe, expect, test } from 'bun:test';
import { createElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  createPptxEditor,
  paintedNotesFindRange,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { FindBar as ReactFindBar, requestFindBar } from '../../react/src/FindBar.tsx';
import { NotesPane as ReactNotesPane } from '../../react/src/panes.tsx';
import { PptxEditorFindBar } from '../src/FindBar.tsx';
import { PptxEditorNotesPane } from '../src/NotesPane.tsx';
import { providePptxEditor } from '../src/composables.ts';

let app: App | null = null;
let vueHost: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  vueHost?.remove();
  app = null;
  vueHost = null;
});

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      { shapes: '', notes: 'slide a slide b slide c' },
      { shapes: '', notes: 'slide d' },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

function mountReact(editor: PptxEditorInstance): HTMLElement {
  const { container } = render(
    createElement(
      PptxEditorContext.Provider,
      { value: editor },
      createElement(
        'div',
        { className: 'pptx-editor' },
        createElement(ReactFindBar),
        createElement(ReactNotesPane)
      )
    )
  );
  return container.firstElementChild as HTMLElement;
}

async function mountVue(editor: PptxEditorInstance): Promise<HTMLElement> {
  vueHost = document.createElement('div');
  vueHost.className = 'pptx-editor';
  document.body.append(vueHost);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => [h(PptxEditorFindBar), h(PptxEditorNotesPane)];
      },
    }) as Component
  );
  app.mount(vueHost);
  await nextTick();
  return vueHost;
}

/** Let React effects and Vue's post-flush and next-tick paint run. */
async function settle(): Promise<void> {
  await act(async () => {
    await nextTick();
    await new Promise((r) => setTimeout(r, 0));
  });
}

async function check(scope: HTMLElement, editor: PptxEditorInstance): Promise<void> {
  act(() => requestFindBar(editor));
  await settle();
  const field = scope.querySelector<HTMLInputElement>('input[aria-label="Find"]')!;
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    setter.call(field, 'slide');
    field.dispatchEvent(new Event('input', { bubbles: true }));
  });
  await settle();
  const count = (): string | null | undefined =>
    scope.querySelector('.pptx-find__count')?.textContent;
  const notes = scope.querySelector<HTMLElement>('.pptx-notes__text')!;
  expect(count()).toBe('1 of 4');
  expect(scope.querySelector('.pptx-find__hint')?.textContent).toBe('In speaker notes');
  expect(paintedNotesFindRange(notes)).toEqual({ start: 0, end: 5 });
  const next = scope.querySelector<HTMLButtonElement>('button[aria-label^="Next match"]')!;
  act(() => next.click());
  await settle();
  expect(count()).toBe('2 of 4');
  expect(paintedNotesFindRange(notes)).toEqual({ start: 8, end: 13 });
  act(() => next.click());
  await settle();
  expect(count()).toBe('3 of 4');
  expect(paintedNotesFindRange(notes)).toEqual({ start: 16, end: 21 });
  // The next match is on slide 2: its notes show the match.
  act(() => next.click());
  await settle();
  expect(count()).toBe('4 of 4');
  expect(editor.snapshot().currentSlide).toBe(1);
  expect(paintedNotesFindRange(notes)).toEqual({ start: 0, end: 5 });
  // A selection on the slide ends the find paint.
  act(() => editor.setSelection({ kind: 'shapes', nodeIds: ['x'] }));
  await settle();
  expect(paintedNotesFindRange(notes) === null).toBe(true);
}

describe('find in the speaker notes', () => {
  test('React: Next steps through the notes matches and paints each one', async () => {
    const editor = await open();
    await check(mountReact(editor), editor);
  });

  test('Vue: Next steps through the notes matches and paints each one', async () => {
    const editor = await open();
    await check(await mountVue(editor), editor);
  });
});
