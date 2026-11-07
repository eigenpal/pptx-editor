// The Animations tab of the side panel: the Vue pane renders the same markup and ARIA as its
// React twin (packages/react/src/AnimationsPane.tsx), adds, edits, and removes effects through
// the engine, and turns the canvas order marks on while it shows.

import { afterEach, describe, expect, test } from 'bun:test';
import { createElement } from 'react';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../scripts/fixtures/timing.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { AnimationsPane as RAnimationsPane } from '../../react/src/AnimationsPane.tsx';
import { AnimationsPane } from '../src/AnimationsPane.tsx';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';

const IN = 914_400;

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

async function mountVue(editor: PptxEditorInstance): Promise<HTMLElement> {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(AnimationsPane);
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

function mountReact(editor: PptxEditorInstance): HTMLElement {
  return render(
    createElement(PptxEditorContext.Provider, { value: editor }, createElement(RAnimationsPane))
  ).container;
}

/** Ids from `useId` differ between the adapters: compare the markup without them. */
const canon = (root: Element): string =>
  canonicalDom(root, { normalizeStyle: true, dropSelectValue: true, formState: true }).replace(
    /\bfor="[^"]*"/g,
    'for=""'
  );

async function deck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const box = (id: number, name: string, x: number) =>
    shape({ id, name, x: x * IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(name)) });
  const bytes = buildPptx({
    slides: [
      { layout: 'blank', shapes: box(2, 'A', 1) + box(3, 'B', 4), ext: clickEntrances([2]) },
      { layout: 'blank', shapes: box(2, 'A', 1) },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const flush = async (): Promise<void> => {
  await nextTick();
  await act(async () => undefined);
};

describe('Animations pane', () => {
  test('Vue renders the React markup, with the effect list and the transition', async () => {
    const editor = await deck();
    const r = canon(mountReact(editor));
    cleanup();
    const v = canon(await mountVue(editor));
    expect(v).toEqual(r);
    expect(v).toContain('pptx-animations__list');
    expect(v).toContain('Appear');
  });

  for (const [name, mount] of [
    ['React', async (e: PptxEditorInstance) => mountReact(e)],
    ['Vue', mountVue],
  ] as const)
    test(`${name}: adds, edits, and removes an effect, and shows the order marks`, async () => {
      const editor = await deck();
      const el = await mount(editor);
      expect(editor.snapshot().animationOrder).toBe(true);
      const b = editor.snapshot().shapes!.find((s) => s.name === 'B')!.nodeId;
      editor.setSelection({ kind: 'shapes', nodeIds: [b] });
      await flush();
      const select = el.querySelectorAll<HTMLSelectElement>('select.pptx-animations__select');
      const presets = select[1]!;
      presets.value = 'spin';
      fireEvent.change(presets);
      await flush();
      const add = [...el.querySelectorAll<HTMLButtonElement>('button.pptx-button')].find(
        (x) => x.textContent === 'Add'
      )!;
      fireEvent.click(add);
      await flush();
      const effects = () => editor.snapshot().animations!.effects;
      expect(effects().map((x) => [x.shapeName, x.preset])).toEqual([
        ['A', 'appear'],
        ['B', 'spin'],
      ]);
      // The new effect is the selected row, with its fields.
      expect(el.querySelector('[aria-pressed="true"]')?.textContent).toContain('Spin');
      const delay = el.querySelector<HTMLInputElement>('input[aria-label="Delay"]')!;
      fireEvent.input(delay, { target: { value: '0.5' } });
      fireEvent.change(delay, { target: { value: '0.5' } });
      fireEvent.blur(delay);
      await flush();
      expect(effects()[1].delayMs).toBe(500);
      fireEvent.click(el.querySelector<HTMLElement>('[aria-label="Move earlier"]')!);
      await flush();
      expect(effects().map((x) => x.shapeName)).toEqual(['B', 'A']);
      fireEvent.click(el.querySelector<HTMLElement>('[aria-label="Remove effect"]')!);
      await flush();
      expect(effects().map((x) => x.shapeName)).toEqual(['A']);
      // The transition menu sets the slide's transition.
      const kind = el.querySelectorAll<HTMLSelectElement>('select.pptx-animations__select')[0]!;
      kind.value = 'push';
      fireEvent.change(kind);
      await flush();
      expect(editor.snapshot().animations!.transition.kind).toBe('push');
      cleanup();
      app?.unmount();
      app = null;
      expect(editor.snapshot().animationOrder).toBe(false);
    });
});
