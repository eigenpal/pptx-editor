// The Chart Data tab of the side panel: the Vue pane renders the same markup and ARIA as its
// React twin (packages/react/src/ChartDataPane.tsx), and applies fields, adds, and removes
// through the engine, one undo step each.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createElement } from 'react';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditorContext } from '../../react/src/context.ts';
import { ChartDataPane as RChartDataPane } from '../../react/src/ChartDataPane.tsx';
import { ChartDataPane } from '../src/ChartDataPane.tsx';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';

const bytes = new Uint8Array(
  readFileSync(join(import.meta.dir, '../../../fixtures/chart-data.pptx'))
);

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

async function deck(slide = 0, select = true): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  editor.goToSlide(slide);
  if (select) {
    const chart = editor.snapshot().shapes!.find((s) => s.name.startsWith('Chart'))!;
    editor.setSelection({ kind: 'shapes', nodeIds: [chart.nodeId] });
  }
  return editor;
}

async function mountVue(editor: PptxEditorInstance): Promise<HTMLElement> {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(ChartDataPane);
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

function mountReact(editor: PptxEditorInstance): HTMLElement {
  return render(
    createElement(PptxEditorContext.Provider, { value: editor }, createElement(RChartDataPane))
  ).container;
}

/** Ids from `useId` differ between the adapters: compare the markup without them. */
const canon = (root: Element): string =>
  canonicalDom(root, { normalizeStyle: true, formState: true }).replace(
    /\b(id|aria-describedby)="[^"]*"/g,
    '$1=""'
  );

const flush = async (): Promise<void> => {
  await nextTick();
  await act(async () => undefined);
};

const data = (editor: PptxEditorInstance) => {
  const r = editor.chartData();
  if (!r.ok) throw new Error(r.reason);
  return r.data;
};

describe('Chart Data tab', () => {
  for (const [what, slide, select] of [
    ['a column chart', 0, true],
    ['a scatter chart', 1, true],
    ['no chart', 0, false],
  ] as const)
    test(`Vue renders the React markup for ${what}`, async () => {
      const r = canon(mountReact(await deck(slide, select)));
      cleanup();
      const v = canon(await mountVue(await deck(slide, select)));
      expect(v).toEqual(r);
    });

  test('Vue renders the React error for text that is not a value', async () => {
    const refuse = async (el: HTMLElement): Promise<string> => {
      const field = el.querySelector<HTMLInputElement>('input[aria-label="Plan, North"]')!;
      fireEvent.input(field, { target: { value: 'many' } });
      fireEvent.change(field, { target: { value: 'many' } });
      fireEvent.blur(field);
      await flush();
      expect(field.getAttribute('aria-invalid')).toBe('true');
      return canon(el);
    };
    const r = await refuse(mountReact(await deck()));
    cleanup();
    const v = await refuse(await mountVue(await deck()));
    expect(v).toEqual(r);
    expect(v).toContain('Enter a number, or leave the cell empty.');
  });

  test('Vue applies a field on Enter as one undo step, and adds and removes', async () => {
    const editor = await deck();
    const el = await mountVue(editor);
    const field = el.querySelector<HTMLInputElement>('input[aria-label="Plan, South"]')!;
    fireEvent.input(field, { target: { value: '77' } });
    fireEvent.keyDown(field, { key: 'Enter' });
    await flush();
    expect(data(editor).series[0]!.values[1]).toBe(77);
    editor.exec({ type: 'undo' });
    await flush();
    expect(data(editor).series[0]!.values[1]).toBe(35);
    expect(field.value).toBe('35');
    const button = (name: string) =>
      [...el.querySelectorAll<HTMLButtonElement>('button')].find(
        (b) => (b.getAttribute('aria-label') ?? b.textContent) === name
      )!;
    fireEvent.click(button('Add series'));
    await flush();
    expect(data(editor).series.map((s) => s.name)).toEqual([
      'Plan',
      'Actual',
      'Forecast',
      'Series 4',
    ]);
    fireEvent.click(button('Remove category South'));
    await flush();
    expect(data(editor).categories.map((c) => c.label)).toEqual(['North', 'East', 'West']);
  });
});
