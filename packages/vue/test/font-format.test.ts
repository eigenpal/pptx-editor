// `useFontFormat`, the Vue twin of React's hook: the same read (null for a field the runs
// disagree about) and the same one-command write.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { useFontFormat, type UseFontFormatReturn } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text: para(run('Hello ', 'sz="1800"') + run('World', 'sz="2400"')),
        }),
      },
    ],
  });

let app: App | null = null;
afterEach(() => {
  app?.unmount();
  app = null;
});

function mount(editor: PptxEditorInstance): UseFontFormatReturn {
  let out: UseFontFormatReturn | null = null;
  const Child = defineComponent({
    setup() {
      out = useFontFormat();
      return () => h('span');
    },
  });
  const Provider = defineComponent({
    setup() {
      providePptxEditor(shallowRef(editor));
      return () => h(Child);
    },
  });
  app = createApp(Provider as Component);
  app.mount(document.createElement('div'));
  return out!;
}

describe('useFontFormat', () => {
  test('reads the mixed selection and applies one command', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck())).ok).toBe(true);
    const hook = mount(editor);
    expect(hook.format.value).toBeNull();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 11 },
    });
    await nextTick();
    expect(hook.isEnabled.value).toBe(true);
    expect(hook.format.value!.disagrees.fontSizePt).toBe(true);
    expect(hook.format.value!.fontSizePt).toBeNull();
    expect(hook.apply({ fontSizePt: 20 })).toBe(true);
    await nextTick();
    expect(hook.format.value!.fontSizePt).toBe(20);
    expect(hook.apply({ fontSizePt: 0 })).toBe(false);
  });
});
