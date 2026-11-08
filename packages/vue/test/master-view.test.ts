// The slide master view in the Vue adapter, the twin of React's `master-view.test.tsx`: the
// slide rail lists masters and layouts, the content shows the bar that names the template and
// closes the view, and the notes pane hides.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContent, PptxEditorNotesPane, PptxEditorSlideRail } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function mounted(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => [h(PptxEditorSlideRail), h(PptxEditorContent), h(PptxEditorNotesPane)];
      },
    })
  );
  app.mount(host);
  await nextTick();
  return editor;
}

async function openMaster(editor: PptxEditorInstance): Promise<void> {
  expect(editor.exec({ type: 'setPresentationView', view: 'slideMaster' }).ok).toBe(true);
  await nextTick();
}

const options = (): HTMLElement[] => [
  ...host!.querySelectorAll<HTMLElement>('.pptx-master-rail__item[role="option"]'),
];

describe('the slide master view in Vue', () => {
  test('the rail lists the master and its layouts with their use', async () => {
    const editor = await mounted();
    await openMaster(editor);
    const rail = host!.querySelector('[role="listbox"][aria-label="Masters and layouts"]')!;
    expect(rail).not.toBeNull();
    const items = options();
    expect(items.length).toBe(editor.snapshot().slides.length);
    expect(items[0]!.dataset.template).toBe('master');
    expect(items[0]!.querySelector('.pptx-master-rail__name')!.textContent).toMatch(/^Master/);
    const current = rail.querySelector<HTMLElement>('[aria-current="true"]')!;
    expect(Number(current.dataset.index)).toBe(editor.snapshot().currentSlide);
    expect(current.textContent).toMatch(/, Used by (1 slide|\d+ slides)$/);
    expect(current.querySelector('.pptx-master-rail__usage')!.textContent).toMatch(/^Used by/);
  });

  test('a click and the arrow keys show another template; Home goes to the master', async () => {
    const editor = await mounted();
    await openMaster(editor);
    options()[2]!.click();
    await nextTick();
    expect(editor.snapshot().currentSlide).toBe(Number(options()[2]!.dataset.index));
    options()[2]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await nextTick();
    expect(editor.snapshot().currentSlide).toBe(Number(options()[3]!.dataset.index));
    options()[3]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    await nextTick();
    expect(editor.snapshot().currentSlide).toBe(Number(options()[0]!.dataset.index));
    const bar = host!.querySelector('.pptx-master-bar[role="region"]')!;
    expect(bar.textContent).toContain('every layout of this master');
  });

  test('the bar closes the view; the notes pane hides while it is open', async () => {
    const editor = await mounted();
    editor.goToSlide(2);
    await nextTick();
    expect(host!.querySelector('.pptx-notes')).not.toBeNull();
    await openMaster(editor);
    expect(host!.querySelector('.pptx-notes')).toBeNull();
    host!.querySelector<HTMLButtonElement>('.pptx-master-bar__close')!.click();
    await nextTick();
    expect(editor.snapshot().presentationView).toBe('normal');
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(host!.querySelector('.pptx-master-bar')).toBeNull();
    // Focus goes to the slide, not to the page, as the bar leaves.
    expect(document.activeElement?.closest('.pptx-canvas')).not.toBeNull();
    expect(host!.querySelector('.pptx-notes')).not.toBeNull();
    // The slide rail is back, on the slide the view opened from.
    const back = host!.querySelector<HTMLElement>('.pptx-rail__item[aria-current="true"]');
    expect(back).not.toBeNull();
    expect(back!.classList.contains('pptx-master-rail__item')).toBe(false);
  });
});
