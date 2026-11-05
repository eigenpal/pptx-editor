// The `slide.*` slots (New slide, Layout, Reset slide, Hide slide) through the Vue
// composables: the same state and the same commands as the React adapter, from the shared
// `chromeCommandState`. Vue has no rail or context menus yet (`deferredInVue`), so a Vue host
// builds that chrome from these composables.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  chromeCommandState,
  createPptxEditor,
  type ChromeSlotId,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import { ToolbarSlot, useEditorCommand, useEditorState, useSlideLayouts } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

function mountWith(editor: PptxEditorInstance, setup: () => () => unknown): HTMLElement {
  // A second mount in one test replaces the first: no app stays in the document for a later
  // file's `document.querySelector` to find.
  app?.unmount();
  host?.remove();
  host = document.createElement('div');
  document.body.append(host);
  const Child = defineComponent({ setup });
  const Provider = defineComponent({
    setup() {
      providePptxEditor(shallowRef(editor));
      return () => h(Child);
    },
  });
  app = createApp(Provider as Component);
  app.mount(host);
  return host;
}

const SLOTS: readonly ChromeSlotId[] = ['slide.add', 'slide.layout', 'slide.reset', 'slide.hide'];

describe('slides slots in Vue', () => {
  test('each slot reports the state the React adapter reads', async () => {
    const editor = await open();
    editor.goToSlide(3);
    const states: Record<string, unknown> = {};
    mountWith(editor, () => {
      const cmds = SLOTS.map((slot) => [slot, useEditorCommand(slot)] as const);
      return () => {
        for (const [slot, cmd] of cmds)
          states[slot] = {
            enabled: cmd.isEnabled.value,
            active: cmd.isActive.value,
            value: cmd.value.value,
          };
        return h('output');
      };
    });
    for (const slot of SLOTS) {
      const shared = chromeCommandState(editor, slot);
      expect(states[slot]).toEqual({
        enabled: shared.isEnabled,
        active: shared.isActive,
        value: shared.value,
      });
    }
    // Slide 4 is hidden, and its layout is Title Only.
    expect(states['slide.hide']).toMatchObject({ enabled: true, active: true });
    const titleOnly = editor.snapshot().layouts!.find((l) => l.name === 'Title Only')!.id;
    expect(states['slide.layout']).toMatchObject({ value: titleOnly });
  });

  test('the commands hide, change the layout, and add a slide on a layout', async () => {
    const editor = await open();
    editor.goToSlide(1);
    const run: Partial<Record<ChromeSlotId, (value?: unknown) => boolean>> = {};
    const el = mountWith(editor, () => {
      const hide = useEditorCommand('slide.hide');
      const layout = useEditorCommand('slide.layout');
      const add = useEditorCommand('slide.add');
      run['slide.hide'] = () => hide.execute();
      run['slide.layout'] = (v) => layout.execute(v);
      run['slide.add'] = (v) => add.execute(v);
      const names = useEditorState((s) => (s.layouts ?? []).map((l) => l.name).join(','));
      return () => h('output', `${hide.isActive.value}|${names.value}`);
    });
    expect(el.textContent).toBe(
      'false|Title and Content,Blank,Title Slide,Two Content,Title Only,Quote'
    );
    run['slide.hide']!();
    await nextTick();
    expect(editor.snapshot().slides[1]!.hidden).toBe(true);
    expect(el.textContent?.startsWith('true|')).toBe(true);
    const ids = Object.fromEntries(editor.snapshot().layouts!.map((l) => [l.name, l.id]));
    // execute says whether the engine ran the command, as in DOCX Editor.
    expect(run['slide.layout']!(ids['Two Content'])).toBe(true);
    expect(editor.snapshot().slides[1]!.layoutId).toBe(ids['Two Content']);
    expect(run['slide.layout']!('no-such-layout')).toBe(false);
    run['slide.add']!(ids['Quote']);
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(editor.snapshot().slides[2]!.layoutId).toBe(ids['Quote']);
  });

  test('view.presentationView switches between Normal and Slide Sorter', async () => {
    const editor = await open();
    let pick: (view: string) => void = () => undefined;
    const el = mountWith(editor, () => {
      const view = useEditorCommand('view.presentationView');
      pick = (v) => view.execute(v);
      const sections = useEditorState((s) => (s.sections ?? []).map((x) => x.name).join(','));
      return () => h('output', `${String(view.value.value)}|${sections.value}`);
    });
    expect(el.textContent).toBe('normal|Opening,Details');
    pick('slideSorter');
    await nextTick();
    expect(el.textContent).toBe('slideSorter|Opening,Details');
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('useSlideLayouts lists the layouts with their master, as in React', async () => {
    const editor = await open();
    const el = mountWith(editor, () => {
      const layouts = useSlideLayouts();
      return () => h('output', layouts.value.map((l) => `${l.name}@${l.masterName}`).join(','));
    });
    expect(el.textContent).toBe(
      'Title and Content@Fixture,Blank@Fixture,Title Slide@Fixture,Two Content@Fixture,Title Only@Fixture,Quote@Fixture'
    );
  });

  test('section.add and section.rename draw their controls, as in React', async () => {
    const editor = await open();
    const el = mountWith(editor, () => () => h(ToolbarSlot, { slotId: 'section.rename' }));
    expect(el.querySelector('[data-slot="section.rename"]') !== null).toBe(true);
    const add = mountWith(editor, () => () => h(ToolbarSlot, { slotId: 'section.add' }));
    expect(add.querySelector('[data-slot="section.add"]') !== null).toBe(true);
  });

  test('a picker slot (the slide layout gallery) draws its control, as in React', async () => {
    const editor = await open();
    const el = mountWith(editor, () => () => h(ToolbarSlot, { slotId: 'slide.layout' }));
    expect(el.querySelector('[data-slot="slide.layout"]') !== null).toBe(true);
  });
});
