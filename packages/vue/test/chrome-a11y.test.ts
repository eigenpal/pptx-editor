// The Vue twin of packages/react/src/__tests__/chrome-a11y.test.tsx: slide items named by
// hidden text with the thumbnail art hidden from assistive technology (the rail and the slide
// sorter), layout tiles named by the layout, and the viewport's "Presentation editor" landmark
// region. The stylesheet checks are shared, so they run in the React test only.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import {
  PptxEditor,
  PptxEditorRoot,
  PptxEditorSlideRail,
  PptxEditorSlideSorter,
  PptxEditorViewport,
} from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';

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
    slides: [1, 2, 3].map((n) => ({
      layout: 'blank' as const,
      hidden: n === 2,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`Text ${n}`)) }),
    })),
  });

async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

async function mount(render: () => unknown, editor?: PptxEditorInstance): Promise<HTMLElement> {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        if (editor) providePptxEditor(shallowRef(editor));
        return render;
      },
    }) as Component
  );
  app.mount(host);
  await tick();
  return host;
}

/** The text a screen reader reads for an element: its content without aria-hidden parts. */
function accessibleText(el: Element): string {
  if (el.getAttribute('aria-hidden') === 'true') return '';
  if (el.hasAttribute('aria-label')) return el.getAttribute('aria-label')!;
  return [...el.childNodes]
    .map((n) =>
      n.nodeType === 3
        ? (n.textContent ?? '')
        : n.nodeType === 1
          ? accessibleText(n as Element)
          : ''
    )
    .join('')
    .trim();
}

function expectNamedItems(selector: string): void {
  const items = [...document.querySelectorAll<HTMLElement>(selector)];
  expect(items.map((i) => i.getAttribute('role'))).toEqual(['option', 'option', 'option']);
  expect(items.map(accessibleText)).toEqual(['Slide 1', 'Slide 2 (hidden)', 'Slide 3']);
  for (const item of items) {
    expect(item.hasAttribute('aria-label')).toBe(false);
    expect(item.querySelector('.pptx-rail__thumb')?.getAttribute('aria-hidden')).toBe('true');
  }
}

describe('slide items', () => {
  test('rail items take their name from hidden text, and the thumbnail art is hidden', async () => {
    const editor = await open();
    await mount(() => h(PptxEditorSlideRail), editor);
    expectNamedItems('.pptx-rail__item');
    expect(document.querySelector('.pptx-rail__number')?.getAttribute('aria-hidden')).toBe('true');
  });

  test('slide sorter tiles take their name from hidden text, and the art is hidden', async () => {
    const editor = await open();
    await mount(() => h(PptxEditorSlideSorter), editor);
    expectNamedItems('.pptx-sorter__item');
    expect(document.querySelector('.pptx-sorter__number')?.getAttribute('aria-hidden')).toBe(
      'true'
    );
  });

  test('a layout tile is named by its layout, not by the placeholder prompts', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    await mount(() => h(PptxEditorSlideRail), editor);
    const add = [...document.querySelectorAll<HTMLElement>('button')].find(
      (b) => b.getAttribute('aria-label') === 'Add slide with layout'
    )!;
    add.click();
    await tick();
    const tiles = [...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].filter((t) =>
      t.classList.contains('pptx-layout-gallery__item')
    );
    expect(accessibleText(tiles[0]!)).toBe('Title and Content');
    expect(
      tiles[0]!.querySelector('.pptx-layout-gallery__thumb')?.getAttribute('aria-hidden')
    ).toBe('true');
  });
});

describe('the editor landmark', () => {
  const region = (): Element | null => document.querySelector('[role="region"].pptx-editor');

  test('the viewport is a region named "Presentation editor"', async () => {
    await mount(() => h(PptxEditorRoot, null, () => h(PptxEditorViewport)));
    expect(region()?.getAttribute('aria-label')).toBe('Presentation editor');
  });

  test('label names the region with text or a translation key, and false removes it', async () => {
    const label = shallowRef<string | false>('Deck preview');
    await mount(() => h(PptxEditorRoot, null, () => h(PptxEditorViewport, { label: label.value })));
    expect(region()?.getAttribute('aria-label')).toBe('Deck preview');
    label.value = 'rail.ariaLabel';
    await tick();
    expect(region()?.getAttribute('aria-label')).toBe('Slides');
    label.value = false;
    await tick();
    const scope = document.querySelector('.pptx-editor')!;
    expect(scope.hasAttribute('role') || scope.hasAttribute('aria-label')).toBe(false);
  });

  test('the default editor puts the toolbar, the rail, the canvas, and the notes in it', async () => {
    let editor: PptxEditorInstance | null = null;
    await mount(() => h(PptxEditor, { onReady: (e: PptxEditorInstance) => (editor = e) }));
    for (let i = 0; i < 300 && !editor; i += 1) await new Promise((r) => setTimeout(r, 10));
    expect((await editor!.load(deck())).ok).toBe(true);
    await tick();
    const r = region()!;
    expect(r.getAttribute('aria-label')).toBe('Presentation editor');
    for (const part of [
      '.pptx-default__dock',
      '.pptx-canvas',
      'nav[aria-label="Slides"], [role="navigation"][aria-label="Slides"]',
      '[aria-label="Speaker notes"]',
    ]) {
      expect(r.querySelector(part) !== null).toBe(true);
    }
  });
});
