// `useBulletFormat`, the Vue twin of React's hook: the same read (undefined when the selected
// paragraphs differ) and the same one-command write.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { useBulletFormat, type UseBulletFormatReturn } from '../src/index.ts';
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
          text:
            para(run('one'), '<a:pPr marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr>') +
            para(run('two')),
        }),
      },
    ],
  });

let app: App | null = null;
afterEach(() => {
  app?.unmount();
  app = null;
});

function mount(editor: PptxEditorInstance): UseBulletFormatReturn {
  let out: UseBulletFormatReturn | null = null;
  const Child = defineComponent({
    setup() {
      out = useBulletFormat();
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

describe('useBulletFormat', () => {
  test('reads a mixed selection as undefined and applies one command', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck())).ok).toBe(true);
    const hook = mount(editor);
    expect(hook.isEnabled.value).toBe(false);
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 1, offset: 1 },
    });
    await nextTick();
    expect(hook.isEnabled.value).toBe(true);
    expect(hook.format.value).toBeUndefined();
    expect(hook.apply({ kind: 'number', scheme: 'romanUcPeriod' })).toBe(true);
    await nextTick();
    expect(hook.format.value).toMatchObject({ kind: 'number', scheme: 'romanUcPeriod' });
    expect(hook.apply(null)).toBe(true);
    await nextTick();
    expect(hook.format.value).toBeNull();
  });
});
