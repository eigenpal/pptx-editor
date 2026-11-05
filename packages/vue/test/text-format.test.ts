// The Vue toolbar draws the script toggles (icon command controls) with the engine's state, and
// runs the same commands as React. The menus and dialogs of the text groups are React only so
// far (scripts/parity/parity.contract.json): the Vue toolbar leaves them out.

import { afterEach, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { attr, child, elements, find, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorToolbar } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';

const IN = 914_400;

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
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('x2')) }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

function mount(editor: PptxEditorInstance): HTMLElement {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        // Off the default bar (the Font dialog sets them): a host adds them back.
        return () =>
          h(PptxEditorToolbar, null, () => [
            h(PptxEditorToolbar.Slot, { slotId: 'script.super' }),
            h(PptxEditorToolbar.Slot, { slotId: 'script.sub' }),
          ]);
      },
    })
  );
  app.mount(host);
  return host;
}

const baseline = (editor: PptxEditorInstance): string | undefined => {
  const part = editor.session!.info.slides[0]!.partName;
  const body = find(xmlPartOf(editor.session!.store.package, part)!.root, 'p:txBody')!;
  const p = elements(body).find((c) => c.local === 'p')!;
  const runs = elements(p).filter((c) => c.local === 'r');
  return attr(child(runs[runs.length - 1], 'a:rPr'), 'baseline');
};

test('superscript and subscript toggle and show their state in the Vue toolbar', async () => {
  const editor = await open();
  const el = mount(editor);
  const button = (slot: string) => el.querySelector<HTMLButtonElement>(`[data-slot="${slot}"]`)!;
  expect(button('script.super').getAttribute('aria-disabled')).toBe('true');
  editor.setSelection({
    kind: 'text',
    nodeId: editor.session!.layout(0)!.items[0]!.element.nodeId,
    anchor: { paragraph: 0, offset: 1 },
    focus: { paragraph: 0, offset: 2 },
  });
  await nextTick();
  button('script.super').click();
  await nextTick();
  expect(baseline(editor)).toBe('30000');
  expect(button('script.super').getAttribute('aria-pressed')).toBe('true');
  button('script.sub').click();
  await nextTick();
  expect(baseline(editor)).toBe('-25000');
  expect(button('script.super').getAttribute('aria-pressed')).toBe('false');
  expect(button('script.sub').getAttribute('aria-pressed')).toBe('true');
  // Menus draw in Vue, as in React. (The Font and Paragraph dialog launchers, Character
  // Spacing, and Columns are not on the preset bar in either adapter.)
  for (const slot of [
    'text.changeCase',
    'paragraph.textDirection',
    'list.bulletStyle',
    'list.numberStyle',
  ])
    expect(el.querySelector(`[data-slot="${slot}"]`) !== null).toBe(true);
});
