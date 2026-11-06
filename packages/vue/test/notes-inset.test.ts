// The default editor keeps its status area and font notice above the notes card: the views box
// carries the card's height in `--pptx-notes-inset`, in the Normal view only. The Vue twin of
// packages/react/src/__tests__/notes-inset.test.tsx.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, h, nextTick, type App } from 'vue';
import { NOTES_INSET_PROPERTY, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { PptxEditor } from '../src/index.ts';

const saved = HTMLElement.prototype.getBoundingClientRect;
let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  HTMLElement.prototype.getBoundingClientRect = saved;
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

async function mountDefault(props: Record<string, unknown> = {}): Promise<HTMLElement> {
  let editor: PptxEditorInstance | null = null;
  host = document.createElement('div');
  document.body.append(host);
  app = createApp({
    render: () => h(PptxEditor, { ...props, onReady: (e: PptxEditorInstance) => (editor = e) }),
  });
  app.mount(host);
  await nextTick();
  await editor!.load(buildPptx({ slides: [{ shapes: '' }] }));
  await nextTick();
  return host.querySelector<HTMLElement>('.pptx-default__views')!;
}

const radio = (name: string) =>
  [...host!.querySelectorAll<HTMLElement>('[role="radio"]')].find(
    (el) => (el.getAttribute('aria-label') ?? el.textContent) === name
  )!;

describe('the default editor and the notes card', () => {
  test('the corner parts float above the notes in the Normal view, at the corner in the sorter', async () => {
    layOut();
    const views = await mountDefault();
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('150px');
    const sorter = radio('Slide sorter');
    sorter.click();
    await nextTick();
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
    // The status area stays mounted, so the view switch keeps its focus.
    expect(radio('Slide sorter')).toBe(sorter);
    radio('Normal').click();
    await nextTick();
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('150px');
  });

  test('without the notes pane the parts stay at the corner', async () => {
    layOut();
    const views = await mountDefault({ notesPane: false });
    expect(views.style.getPropertyValue(NOTES_INSET_PROPERTY)).toBe('0px');
  });
});
