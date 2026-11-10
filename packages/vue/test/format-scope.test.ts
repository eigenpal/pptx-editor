import { afterEach, expect, test } from 'bun:test';
import { cleanup } from '@testing-library/react';
import { createElement } from 'react';
import { h, type App } from 'vue';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { Toolbar } from '../../react/src/index.ts';
import { FormatPanel as ReactPanel } from '../../react/src/FormatPanel.tsx';
import { PptxEditorToolbar } from '../src/index.ts';
import { FormatPanel as VuePanel } from '../src/FormatPanel.tsx';
import { mountVue, openDeck, renderReact, selectText } from './toolbar-dom.ts';

let editor: PptxEditorInstance;
let mounted: { host: HTMLElement; app: App } | undefined;
afterEach(() => {
  cleanup();
  mounted?.app.unmount();
  mounted?.host.remove();
  mounted = undefined;
  editor?.destroy();
});

for (const adapter of ['react', 'vue'] as const) {
  for (const custom of [false, true]) {
    test(`${adapter}: inspector scopes controls independently of toolbar overrides (${custom})`, async () => {
      editor = await openDeck();
      selectText(editor);
      let host: HTMLElement;
      if (adapter === 'react') {
        host = renderReact(
          editor,
          createElement(
            'div',
            null,
            createElement(
              Toolbar,
              null,
              ...(custom
                ? [
                    createElement(Toolbar.Slot, { slotId: 'font.family', hidden: true }),
                    createElement(
                      Toolbar.Slot,
                      { slotId: 'text.bold' },
                      createElement('button', { 'data-host-bold': '' }, 'Host bold')
                    ),
                  ]
                : [])
            ),
            createElement(ReactPanel)
          )
        );
      } else {
        mounted = await mountVue(editor, () =>
          h('div', [
            h(PptxEditorToolbar, null, () =>
              custom
                ? [
                    h(PptxEditorToolbar.Slot, { slotId: 'font.family', hidden: true }),
                    h(PptxEditorToolbar.Slot, { slotId: 'text.bold' }, () =>
                      h('button', { 'data-host-bold': '' }, 'Host bold')
                    ),
                  ]
                : []
            ),
            h(VuePanel),
          ])
        );
        host = mounted.host;
      }
      const toolbar = host.querySelector('.pptx-toolbar')!;
      const panel = host.querySelector('.pptx-format-panel')!;
      expect(toolbar.querySelectorAll('[data-slot="font.family"]')).toHaveLength(custom ? 0 : 1);
      expect(toolbar.querySelectorAll('[data-host-bold]')).toHaveLength(custom ? 1 : 0);
      expect(panel.querySelectorAll('[data-host-bold]')).toHaveLength(0);
      expect(panel.querySelectorAll('[data-format-slot="font.family"]')).toHaveLength(1);
      expect(panel.querySelectorAll('[data-format-slot="font.family"] button')).toHaveLength(1);
      expect(
        panel.querySelectorAll('[data-format-slot="text.bold"] [data-slot="text.bold"]')
      ).toHaveLength(1);
    });
  }
}
