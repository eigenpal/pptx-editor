// The Diagram tab of the side panel: the Vue pane renders the same markup and ARIA as its React
// twin (packages/react/src/DiagramTextPane.tsx), and applies a field, an add, and a delete
// through the engine, one undo step each.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createElement } from 'react';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditorContext } from '../../react/src/context.ts';
import { DiagramTextPane as RDiagramTextPane } from '../../react/src/DiagramTextPane.tsx';
import { DiagramTextPane } from '../src/DiagramTextPane.tsx';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';

const bytes = new Uint8Array(
  readFileSync(join(import.meta.dir, '../../../fixtures/diagram-editing.pptx'))
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
    const frame = editor.snapshot().shapes!.find((s) => s.name.startsWith('Diagram'))!;
    editor.setSelection({ kind: 'shapes', nodeIds: [frame.nodeId] });
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
        return () => h(DiagramTextPane);
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

function mountReact(editor: PptxEditorInstance): HTMLElement {
  return render(
    createElement(PptxEditorContext.Provider, { value: editor }, createElement(RDiagramTextPane))
  ).container;
}

/**
 * Ids from `useId` differ between the adapters: compare the markup without them. React writes
 * a textarea's value as its text and Vue as a property, so each textarea shows its live value.
 */
const canon = (root: Element): string =>
  canonicalDom(root, { normalizeStyle: true, formState: true })
    .replace(/\b(id|aria-describedby)="[^"]*"/g, '$1=""')
    .replace(/<textarea ([^>]*?)( value="[^"]*")?>[^<]*<\/textarea>/g, '<textarea $1></textarea>') +
  JSON.stringify([...root.querySelectorAll('textarea')].map((t) => t.value));

const flush = async (): Promise<void> => {
  await nextTick();
  await act(async () => undefined);
};

const texts = (editor: PptxEditorInstance): string[] =>
  editor.snapshot().diagram!.nodes.map((n) => n.text);

describe('Diagram tab', () => {
  for (const [what, slide, select] of [
    ['a process', 0, true],
    ['a vertical list', 1, true],
    ['no diagram', 0, false],
  ] as const)
    test(`Vue renders the React markup for ${what}`, async () => {
      const r = canon(mountReact(await deck(slide, select)));
      cleanup();
      const v = canon(await mountVue(await deck(slide, select)));
      expect(v).toEqual(r);
    });

  test('Vue applies a field on Enter as one undo step, and adds and deletes', async () => {
    const editor = await deck();
    const before = texts(editor);
    const el = await mountVue(editor);
    const field = el.querySelector<HTMLTextAreaElement>('textarea[aria-label="Text of shape 1"]')!;
    fireEvent.input(field, { target: { value: 'Scope' } });
    fireEvent.keyDown(field, { key: 'Enter' });
    await flush();
    expect(texts(editor)[0]).toBe('Scope');
    editor.exec({ type: 'undo' });
    await flush();
    expect(field.value).toBe(before[0]!);
    const button = (name: string) =>
      [...el.querySelectorAll<HTMLButtonElement>('button')].find(
        (b) => (b.getAttribute('aria-label') ?? b.textContent) === name
      )!;
    fireEvent.click(button(`Add shape after ${before[0]}`));
    await flush();
    expect(texts(editor)).toEqual([before[0]!, '', ...before.slice(1)]);
    await flush();
    expect(document.activeElement?.getAttribute('aria-label')).toBe('Text of shape 2');
    fireEvent.click(button('Delete shape Empty shape'));
    await flush();
    expect(texts(editor)).toEqual(before);
    // After a delete, the focus goes to the field that took the deleted row's place.
    expect(document.activeElement?.getAttribute('aria-label')).toBe('Text of shape 2');
  });
});
