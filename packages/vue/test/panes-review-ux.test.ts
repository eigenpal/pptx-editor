// The Vue twin of packages/react/src/__tests__/panes-review-ux.test.tsx: a quick drag in the
// rail or the sorter reorders before Vue renders again, the zoom trigger says whether its menu
// is open, the side panel landmark has a name, and focus goes to the slide when the sorter
// closes with the rail hidden.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import {
  PptxEditorSidePanel,
  PptxEditorSlideRail,
  PptxEditorSlideSorter,
  PptxEditorZoomPill,
} from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { DefaultWorkspace } from '../src/SlideSorter.tsx';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

const IN = 914_400;
const tick = async (): Promise<void> => {
  await nextTick();
  await nextTick();
};
const deck = () =>
  buildPptx({
    slides: [1, 2, 3].map((i) => ({
      layout: 'blank' as const,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`S${i}`)) }),
    })),
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

async function mount(render: () => unknown, editor: PptxEditorInstance): Promise<HTMLElement> {
  host = document.createElement('div');
  host.className = 'pptx-editor';
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return render;
      },
    }) as Component
  );
  app.mount(host);
  await tick();
  return host;
}

const drag = (type: string, init: Record<string, number> = {}): Event =>
  Object.assign(new Event(type, { bubbles: true, cancelable: true }), init);

/** Drag item 0 onto item 2 with no render between the events, as a quick drag does. */
function quickDrag(items: HTMLElement[]): boolean {
  items[0]!.dispatchEvent(drag('dragstart'));
  const over = drag('dragover', { clientX: 1000, clientY: 1000 });
  items[2]!.dispatchEvent(over);
  items[2]!.dispatchEvent(drag('drop'));
  return over.defaultPrevented;
}

const ids = (editor: PptxEditorInstance) => editor.snapshot().slides.map((s) => s.id);

describe('a quick drag', () => {
  test('in the rail: dragover takes the drop before a render, and the drop moves the slide', async () => {
    const editor = await open();
    const first = ids(editor)[0]!;
    const el = await mount(() => h(PptxEditorSlideRail), editor);
    // A dragover that does not call preventDefault() refuses the drop.
    expect(quickDrag([...el.querySelectorAll<HTMLElement>('.pptx-rail__item')])).toBe(true);
    expect(ids(editor).indexOf(first)).toBe(2);
  });

  test('in the sorter: the same', async () => {
    const editor = await open();
    const first = ids(editor)[0]!;
    const el = await mount(() => h(PptxEditorSlideSorter), editor);
    expect(quickDrag([...el.querySelectorAll<HTMLElement>('.pptx-sorter__item')])).toBe(true);
    expect(ids(editor).indexOf(first)).toBe(2);
  });
});

describe('chrome names and states', () => {
  test('the zoom trigger says whether its menu is open', async () => {
    const editor = await open();
    const el = await mount(() => h(PptxEditorZoomPill), editor);
    const trigger = el.querySelector<HTMLElement>('.pptx-zoom__level')!;
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    trigger.click();
    await tick();
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    trigger.click();
    await tick();
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
  });

  test('the side panel landmark has a name', async () => {
    const editor = await open();
    const el = await mount(() => h(PptxEditorSidePanel), editor);
    expect(el.querySelector('aside')!.getAttribute('aria-label')).toBe('Side panel');
  });
});

describe('leaving the sorter', () => {
  test('with the rail hidden, focus goes to the slide, not to the page body', async () => {
    const editor = await open();
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    const el = await mount(
      () =>
        h(DefaultWorkspace, null, {
          normal: () => [h(PptxEditorSlideRail), h('div', { class: 'pptx-surface', tabindex: 0 })],
        }),
      editor
    );
    // A narrow window hides the rail: its items cannot take focus.
    const block = HTMLElement.prototype.focus;
    HTMLElement.prototype.focus = function (this: HTMLElement, o?: FocusOptions) {
      if (!this.closest('.pptx-rail')) block.call(this, o);
    };
    try {
      el.querySelector<HTMLElement>('.pptx-sorter__item')!.focus();
      editor.exec({ type: 'setPresentationView', view: 'normal' });
      await tick();
      expect(document.activeElement === el.querySelector('.pptx-surface')).toBe(true);
    } finally {
      HTMLElement.prototype.focus = block;
    }
  });
});
