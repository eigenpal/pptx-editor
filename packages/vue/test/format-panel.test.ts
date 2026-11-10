import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createElement } from 'react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { installDrawingRenderers } from '@pptx-editor.dev/core/layout';
import { buildShowcase } from '../../../scripts/sample-deck/showcase.ts';
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

function choose(slide: number, name: string) {
  editor.goToSlide(slide);
  const item = editor.session!.layout(slide)!.items.find((i) => i.element.name === name)!;
  editor.setSelection({ kind: 'shapes', nodeIds: [item.element.nodeId] });
}
const field = (id: string) =>
  host.querySelector<HTMLInputElement>(`input[data-format-property="${id}"]`)!;

for (const adapter of ['react', 'vue'] as const)
  describe(`${adapter} format inspector`, () => {
    async function mount(custom = false) {
      editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
      expect((await editor.load(buildShowcase())).ok).toBe(true);
      choose(6, 'Three stops');
      if (adapter === 'react') {
        host = render(
          createElement(
            PptxEditorContext.Provider,
            { value: editor },
            createElement(
              ReactPanel,
              custom
                ? {
                    sections: ['effects', 'fill'],
                    sectionContent: { fill: createElement('p', null, 'Custom fill') },
                  }
                : {}
            )
          )
        ).container;
      } else {
        host = document.createElement('div');
        document.body.append(host);
        app = createApp(
          defineComponent({
            setup() {
              providePptxEditor(shallowRef(editor));
              return () =>
                h(
                  VuePanel,
                  custom ? { sections: ['effects', 'fill'] } : {},
                  custom ? { fill: () => h('p', 'Custom fill') } : {}
                );
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
    async function input(id: string, value: string) {
      await update(() => {
        if (adapter === 'react') fireEvent.change(field(id), { target: { value } });
        else {
          field(id).value = value;
          field(id).dispatchEvent(new Event('input', { bubbles: true }));
        }
      });
    }
    async function selectStop(index: number) {
      const select = host.querySelector<HTMLSelectElement>('[data-format-stop-select]')!;
      await update(() =>
        fireEvent.change(select, { target: { value: select.options[index]!.value } })
      );
    }
    test('edits gradient details and preserves theme stops through undo', async () => {
      await mount();
      expect(field('gradient.angle').value).toBe('45');
      await selectStop(1);
      expect(field('gradient-stop.1.position').value).toBe('50');
      const before = editor.shapeFormat();
      await input('gradient.angle', '75');
      await update(() => fireEvent.keyDown(field('gradient.angle'), { key: 'Enter' }));
      const after = editor.shapeFormat();
      expect(after.ok && after.format.gradient?.angle).toBe(75);
      expect(after.ok && after.format.gradient?.stops).toEqual(
        before.ok && before.format.gradient?.stops
      );
      await update(() => {
        editor.exec({ type: 'undo' });
      });
      expect(field('gradient.angle').value).toBe('45');
    });
    test('drops drafts on selection changes and edits a custom shadow', async () => {
      await mount();
      await input('gradient.angle', '90');
      await update(() => choose(8, 'Outer shadow'));
      expect(field('gradient.angle')).toBeNull();
      expect(field('shadow.distance').value).toBe('8');
      await input('shadow.distance', '12');
      await update(() => fireEvent.blur(field('shadow.distance')));
      expect(editor.shapeFormat().ok && field('shadow.distance').value).toBe('12');
      await update(() => choose(6, 'Three stops'));
      expect(field('gradient.angle').value).toBe('45');
    });
    test('keeps invalid drafts with errors and respects view mode', async () => {
      await mount();
      await input('gradient-stop.0.opacity', '101');
      await update(() => fireEvent.keyDown(field('gradient-stop.0.opacity'), { key: 'Enter' }));
      expect(field('gradient-stop.0.opacity').getAttribute('aria-invalid')).toBe('true');
      await update(() => fireEvent.keyDown(field('gradient-stop.0.opacity'), { key: 'Escape' }));
      expect(field('gradient-stop.0.opacity').value).toBe('100');
      await update(() => editor.setMode('view'));
      expect(field('gradient.angle').disabled).toBe(true);
    });
    test('shows only common formatting for a mixed picture and shape selection', async () => {
      await mount();
      await update(() => {
        editor.goToSlide(9);
        const items = editor.session!.layout(9)!.items;
        const picture = items.find((i) => i.element.kind === 'picture')!;
        const shape = items.find((i) => i.element.kind === 'shape')!;
        editor.setSelection({
          kind: 'shapes',
          nodeIds: [picture.element.nodeId, shape.element.nodeId],
        });
      });
      expect(host.querySelector('[data-format-section="fill"]')).toBeNull();
      expect(host.querySelector('[data-format-section="text"]')).toBeNull();
      expect(host.querySelector('[data-format-section="line"]')).not.toBeNull();
    });
    test('keeps a gradient stop draft bound to its source after another stop is removed', async () => {
      await mount();
      await selectStop(1);
      await input('gradient-stop.1.opacity', '65');
      await update(() => {
        const read = editor.shapeFormat();
        if (!read.ok) throw new Error(read.reason);
        editor.exec({
          type: 'patchShapeFormat',
          target: read.format.target,
          patch: { kind: 'gradient', removeStop: read.format.gradient!.stops[0]!.id },
        });
      });
      expect(field('gradient-stop.0.opacity').value).toBe('65');
      await update(() => fireEvent.keyDown(field('gradient-stop.0.opacity'), { key: 'Enter' }));
      const read = editor.shapeFormat();
      expect(read.ok && read.format.gradient!.stops[0]!.opacity).toBe(0.65);
      expect(read.ok && read.format.gradient!.stops[1]!.opacity).toBe(1);
    });
    test('cancels a draft when its source property changes', async () => {
      await mount();
      await input('gradient.angle', '90');
      await update(() => {
        const read = editor.shapeFormat();
        if (!read.ok) throw new Error(read.reason);
        editor.exec({
          type: 'patchShapeFormat',
          target: read.format.target,
          patch: { kind: 'gradient', angle: 15 },
        });
      });
      expect(field('gradient.angle').value).toBe('15');
      await update(() => fireEvent.blur(field('gradient.angle')));
      const read = editor.shapeFormat();
      expect(read.ok && read.format.gradient!.angle).toBe(15);
    });
    test('updates numeric effect availability after a drawing module install', async () => {
      await mount();
      await update(() => choose(8, 'Inner shadow'));
      expect(field('shadow.distance').disabled).toBe(true);
      try {
        await update(() =>
          installDrawingRenderers({
            effects: { body: () => null, glow: () => null, reflection: () => null },
          })
        );
        expect(field('shadow.distance').disabled).toBe(false);
      } finally {
        await update(() => installDrawingRenderers({ effects: undefined }));
      }
    });
    test('adds a stop and moves focus to its active editor', async () => {
      await mount();
      const add = [...host.querySelectorAll('button')].find((b) => b.textContent === 'Add')!;
      await update(() => fireEvent.click(add));
      const select = host.querySelector<HTMLSelectElement>('[data-format-stop-select]')!;
      expect(select.options.length).toBe(4);
      expect(select.selectedIndex).toBe(2);
      expect(document.activeElement).toBe(select);
      expect(field('gradient-stop.2.position').value).toBe('50');
    });
    test('supports stop keyboard selection and focus after removal', async () => {
      await mount();
      const markers = host.querySelectorAll<HTMLButtonElement>('[data-format-stop-marker]');
      await update(() => fireEvent.keyDown(markers[0]!, { key: 'ArrowRight' }));
      expect(markers[1]!.getAttribute('aria-pressed')).toBe('true');
      expect(document.activeElement).toBe(markers[1]);
      const remove = host.querySelector('[data-format-stop] > button')!;
      await update(() => fireEvent.click(remove));
      const select = host.querySelector<HTMLSelectElement>('[data-format-stop-select]')!;
      expect(select.options.length).toBe(2);
      expect(select.selectedIndex).toBe(1);
      expect(document.activeElement).toBe(select);
    });
    test('keeps numeric drafts during composition', async () => {
      await mount();
      await input('gradient.angle', '72');
      await update(() =>
        fireEvent.keyDown(field('gradient.angle'), { key: 'Enter', isComposing: true })
      );
      const before = editor.shapeFormat();
      expect(before.ok && before.format.gradient?.angle).toBe(45);
      expect(field('gradient.angle').value).toBe('72');
      await update(() => fireEvent.keyDown(field('gradient.angle'), { key: 'Enter' }));
      const after = editor.shapeFormat();
      expect(after.ok && after.format.gradient?.angle).toBe(72);
    });
    test('validates exact line width and geometry without losing invalid input', async () => {
      await mount();
      await input('line.width', '0');
      await update(() => fireEvent.keyDown(field('line.width'), { key: 'Enter' }));
      expect(field('line.width').getAttribute('aria-invalid')).toBe('true');
      expect(host.textContent).toContain('greater than 0');
      await input('line.width', '0.125');
      await update(() => fireEvent.keyDown(field('line.width'), { key: 'Enter' }));
      expect(Number(field('line.width').value)).toBe(Math.round(0.125 * 12700) / 12700);
      await input('geometry.width', '-1');
      await update(() => fireEvent.keyDown(field('geometry.width'), { key: 'Enter' }));
      expect(field('geometry.width').getAttribute('aria-invalid')).toBe('true');
      expect(field('geometry.width').value).toBe('-1');
      await update(() => fireEvent.keyDown(field('geometry.width'), { key: 'Escape' }));
      expect(field('geometry.width').getAttribute('aria-invalid')).not.toBe('true');
    });
    test('persists disclosure choices when changing selections in the same category', async () => {
      await mount();
      await update(() => fireEvent.click(host.querySelector('[data-format-toggle="geometry"]')!));
      expect(
        host.querySelector<HTMLDetailsElement>('[data-format-section="geometry"] details')!.open
      ).toBe(true);
      await update(() => choose(6, 'Linear'));
      expect(
        host.querySelector<HTMLDetailsElement>('[data-format-section="geometry"] details')!.open
      ).toBe(true);
    });
    test('orders and replaces sections without rendering omitted sections', async () => {
      await mount(true);
      const sections = [...host.querySelector('.pptx-format-sections')!.children];
      expect(sections.map((s) => s.getAttribute('data-format-section'))).toEqual([
        'effects',
        'fill',
      ]);
      expect(host.textContent).toContain('Custom fill');
      expect(field('gradient.angle')).toBeNull();
      expect(host.querySelector('[data-format-section="geometry"]')).toBeNull();
    });
  });
