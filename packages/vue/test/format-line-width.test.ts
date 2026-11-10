import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createElement } from 'react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, shape } from '../../../scripts/fixtures/builder.ts';
import { FormatLineWidth as ReactWidth } from '../../react/src/FormatLineWidth.tsx';
import { PptxEditorContext } from '../../react/src/context.ts';
import { FormatLineWidth as VueWidth } from '../src/FormatLineWidth.tsx';
import { providePptxEditor } from '../src/composables.ts';
let app: App | undefined;
let editor: PptxEditorInstance;
let host: HTMLElement;
afterEach(() => {
  cleanup();
  app?.unmount();
  app = undefined;
  host?.remove();
  editor?.destroy();
});
const source = () =>
  serializeXmlPart(xmlPartOf(editor.session!.store.package, '/ppt/slides/slide1.xml')!);
for (const adapter of ['react', 'vue'] as const)
  describe(`${adapter} concise line width`, () => {
    async function update(run: () => void) {
      if (adapter === 'react') act(run);
      else {
        run();
        await nextTick();
      }
    }
    async function mount(multiple = false, firstWidth = 39687) {
      editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
      await editor.load(
        buildPptx({
          slides: [
            {
              layout: 'blank',
              shapes: [firstWidth, 39686]
                .map((width, index) =>
                  shape({
                    id: index + 2,
                    name: `Line ${index}`,
                    prst: 'line',
                    x: 0,
                    y: index * 914400,
                    cx: 914400,
                    cy: 0,
                    line: `<a:ln w="${width}"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>`,
                  })
                )
                .join(''),
            },
          ],
        })
      );
      const ids = [...editor.snapshot().shapes!]
        .sort((a, b) => a.name.localeCompare(b.name))
        .map((s) => s.nodeId);
      editor.setSelection({ kind: 'shapes', nodeIds: multiple ? ids : ids.slice(0, 1) });
      if (adapter === 'react')
        host = render(
          createElement(PptxEditorContext.Provider, { value: editor }, createElement(ReactWidth))
        ).container;
      else {
        host = document.createElement('div');
        document.body.append(host);
        app = createApp(
          defineComponent({
            setup() {
              providePptxEditor(shallowRef(editor));
              return () => h(VueWidth);
            },
          })
        );
        app.mount(host);
        await nextTick();
      }
    }
    const input = () => host.querySelector<HTMLInputElement>('input')!;
    async function type(value: string) {
      await update(() => {
        if (adapter === 'react') fireEvent.change(input(), { target: { value } });
        else {
          input().value = value;
          fireEvent.input(input());
        }
      });
    }
    test('rounded display preserves imported precision on untouched blur and equivalent values', async () => {
      await mount();
      const before = source();
      expect(input().value).toBe('3.12');
      await update(() => {
        fireEvent.focus(input());
        fireEvent.blur(input());
      });
      expect(source()).toBe(before);
      for (const value of ['3.12', String(39687 / 12700)]) {
        await type(value);
        await update(() => fireEvent.keyDown(input(), { key: 'Enter' }));
        expect(source()).toBe(before);
        expect(input().value).toBe('3.12');
      }
      await type('3.125');
      await update(() => fireEvent.keyDown(input(), { key: 'Enter' }));
      expect(source()).toBe(before.replace('w="39687"', 'w="39688"'));
      expect(input().value).toBe('3.13');
      await update(() => {
        editor.exec({ type: 'undo' });
      });
      expect(source()).toBe(before);
      expect(input().value).toBe('3.12');
    });
    test('10000 EMU displays 0.79pt without changing its stored width', async () => {
      await mount(false, 10000);
      const before = source();
      expect(input().value).toBe('0.79');
      expect(editor.snapshot().geometry!.lineWidth).toBe(0.79);
      await update(() => {
        fireEvent.focus(input());
        fireEvent.blur(input());
      });
      expect(source()).toBe(before);
      await type('0.79');
      await update(() => fireEvent.keyDown(input(), { key: 'Enter' }));
      expect(source()).toBe(before);
    });
    test('different exact widths stay mixed even when they round to the same display', async () => {
      await mount(true);
      const before = source();
      expect(input().value).toBe('');
      expect(input().placeholder).toBe('Mixed');
      await update(() => {
        fireEvent.focus(input());
        fireEvent.blur(input());
      });
      expect(source()).toBe(before);
    });
  });
