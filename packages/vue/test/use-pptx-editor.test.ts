// `usePptxEditor()` outside `PptxEditorRoot` returns a null ref with a development warning, as
// the other parts do (`warnOutsideRoot`); under a root or a host's `providePptxEditor` it stays
// quiet. The twin of the React test.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { providePptxEditor, usePptxEditor } from '../src/composables.ts';
import { PptxEditorRoot } from '../src/index.ts';

const OUTSIDE = 'usePptxEditor is used outside PptxEditor.Root';
const apps: [App, HTMLElement][] = [];
afterEach(() => {
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
});

function mount(render: () => ReturnType<typeof h>): void {
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({ render });
  app.mount(host);
  apps.push([app, host]);
}

const seen: (PptxEditorInstance | null)[] = [];
const Probe = defineComponent({
  setup() {
    const editor = usePptxEditor();
    return () => {
      seen.push(editor.value);
      return null;
    };
  },
});

const warned = (warn: { mock: { calls: unknown[][] } }): boolean =>
  warn.mock.calls.some((c) => String(c[0]).includes(OUTSIDE));

describe('usePptxEditor (Vue)', () => {
  test('outside the root it returns null and warns', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    mount(() => h(Probe));
    await nextTick();
    expect(seen.at(-1) === null).toBe(true);
    expect(warned(warn)).toBe(true);
    warn.mockRestore();
  });

  test('inside the root, or with a host provider, it does not warn', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    mount(() => h(PptxEditorRoot, null, { default: () => h(Probe) }));
    await nextTick();
    const editor = createPptxEditor({ fixedMeasurer: true });
    mount(() =>
      h(
        defineComponent({
          setup() {
            providePptxEditor(shallowRef(editor));
            return () => h(Probe);
          },
        })
      )
    );
    await nextTick();
    expect(seen.includes(editor)).toBe(true);
    expect(warned(warn)).toBe(false);
    warn.mockRestore();
    editor.destroy();
  });
});
