import { buildPptx, shape } from '../../../scripts/fixtures/builder.ts';
import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createElement } from 'react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { FormatPanel as ReactPanel } from '../../react/src/FormatPanel.tsx';
import { PptxEditorContext } from '../../react/src/context.ts';
import { FormatPanel as VuePanel } from '../src/FormatPanel.tsx';
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

const box = { x: 0, y: 0, cx: 914400, cy: 914400 };
const deck = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Half',
            ...box,
            fill: '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill>',
          }) +
          shape({
            id: 3,
            name: 'Gradient',
            ...box,
            fill: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>',
          }) +
          shape({ id: 4, name: 'Empty', ...box, fill: '<a:noFill/>' }),
      },
    ],
  });

const idOf = (name: string) =>
  editor.session!.layout(0)!.items.find((i) => i.element.name === name)!.element.nodeId;
const opacity = () => {
  const read = editor.shapeFormat();
  return read.ok ? read.format.solid?.opacity : undefined;
};
const number = () =>
  host.querySelector<HTMLInputElement>('input[data-format-property="fill.transparency"]');
const slider = () =>
  host.querySelector<HTMLInputElement>(
    'input[type="range"][data-format-slider="fill.transparency"]'
  );

for (const adapter of ['react', 'vue'] as const)
  describe(`${adapter} fill transparency`, () => {
    async function mount(...names: string[]) {
      editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
      expect((await editor.load(deck())).ok).toBe(true);
      editor.setSelection({ kind: 'shapes', nodeIds: names.map(idOf) });
      if (adapter === 'react') {
        host = render(
          createElement(
            PptxEditorContext.Provider,
            { value: editor },
            createElement(ReactPanel, {})
          )
        ).container;
      } else {
        host = document.createElement('div');
        document.body.append(host);
        app = createApp(
          defineComponent({
            setup() {
              providePptxEditor(shallowRef(editor));
              return () => h(VuePanel);
            },
          })
        );
        app.mount(host);
        await nextTick();
      }
    }
    async function update(fn: () => void) {
      if (adapter === 'react') act(fn);
      else {
        fn();
        await nextTick();
      }
    }
    /** One step of a drag or an arrow key: the value moves, then `change` fires on release. */
    async function slide(value: string, release: boolean) {
      await update(() => {
        if (adapter === 'react') fireEvent.input(slider()!, { target: { value } });
        else {
          slider()!.value = value;
          slider()!.dispatchEvent(new Event('input', { bubbles: true }));
        }
      });
      if (release)
        await update(() => {
          slider()!.dispatchEvent(new Event('change', { bubbles: true }));
        });
    }

    test('shows the transparency of a solid fill under Fill color', async () => {
      await mount('Half');
      const control = host.querySelector('[data-format-fill-transparency]');
      expect(control?.previousElementSibling?.getAttribute('data-format-slot')).toBe('shape.fill');
      expect(number()?.value).toBe('50');
      expect(slider()?.value).toBe('50');
      expect(slider()?.getAttribute('aria-label')).toBe('Transparency');
      expect(slider()?.getAttribute('aria-valuetext')).toBe('50%');
      expect(slider()?.min).toBe('0');
      expect(slider()?.max).toBe('100');
    });

    test('writes a typed percent and undoes it in one step', async () => {
      await mount('Half');
      await update(() => {
        if (adapter === 'react') fireEvent.change(number()!, { target: { value: '30' } });
        else {
          number()!.value = '30';
          number()!.dispatchEvent(new Event('input', { bubbles: true }));
        }
      });
      await update(() => fireEvent.keyDown(number()!, { key: 'Enter' }));
      expect(opacity()).toBe(0.7);
      expect(slider()?.value).toBe('30');
      await update(() => {
        editor.exec({ type: 'undo' });
      });
      expect(opacity()).toBe(0.5);
      expect(number()?.value).toBe('50');
    });

    test('writes a slider drag once, on release', async () => {
      await mount('Half');
      await slide('40', false);
      await slide('20', false);
      expect(opacity()).toBe(0.5);
      expect(number()?.value).toBe('20');
      expect(slider()?.getAttribute('aria-valuetext')).toBe('20%');
      await slide('20', true);
      expect(opacity()).toBe(0.8);
      await update(() => {
        editor.exec({ type: 'undo' });
      });
      expect(opacity()).toBe(0.5);
      expect(slider()?.value).toBe('50');
    });

    test('is hidden for gradient fills, no fill, and several shapes', async () => {
      for (const names of [['Gradient'], ['Empty'], ['Half', 'Gradient']]) {
        await mount(...names);
        expect(host.querySelector('[data-format-fill-transparency]')).toBeNull();
        expect(slider()).toBeNull();
        cleanup();
        app?.unmount();
        app = undefined;
        host.remove();
        editor.destroy();
      }
    });

    test('is disabled in view mode', async () => {
      await mount('Half');
      await update(() => editor.setMode('view'));
      expect(slider()?.disabled).toBe(true);
      expect(number()?.disabled).toBe(true);
    });
  });
