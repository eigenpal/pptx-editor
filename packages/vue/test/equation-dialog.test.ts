import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  createPptxEditor,
  toolbarCommandState,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { providePptxEditor } from '../src/composables.ts';
import { EquationDialog } from '../src/index.ts';

// The Vue twin of packages/react/src/__tests__/equation-dialog.test.tsx: a gallery pick inserts
// an equation, the linear form of a selected equation edits it as one undo step, and a linear
// form the editor cannot read blocks OK.

const FIXTURE = join(import.meta.dir, '../../../fixtures/equations.pptx');

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(new Uint8Array(readFileSync(FIXTURE)))).ok).toBe(true);
  editor.selectSlide(1);
  return editor;
}

async function show(editor: PptxEditorInstance): Promise<{ closed: number }> {
  const state = { closed: 0 };
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () =>
          h(EquationDialog as Component, { open: true, onClose: () => (state.closed += 1) });
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return state;
}

const field = (): HTMLInputElement =>
  document.querySelector<HTMLInputElement>('[data-pptx-field="linear"] input')!;
const apply = (): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>('[data-pptx-part="apply"]')!;
const title = (): string => document.querySelector('[data-pptx-part="title"]')!.textContent ?? '';

async function type(value: string): Promise<void> {
  field().value = value;
  field().dispatchEvent(new Event('input', { bubbles: true }));
  await nextTick();
}

describe('Equation dialog', () => {
  test('a gallery pick fills the linear form, and Insert adds the equation', async () => {
    const editor = await open();
    expect(toolbarCommandState(editor, 'insert.equation')).toMatchObject({ enabled: true });
    const state = await show(editor);
    expect(title()).toBe('Insert equation');
    const pick = [...document.querySelectorAll<HTMLButtonElement>('button')].find((b) =>
      b.textContent?.startsWith('Area of a circle')
    )!;
    pick.click();
    await nextTick();
    expect(field().value).toBe('A=\\pi r^2');
    apply().click();
    await nextTick();
    expect(state.closed).toBe(1);
    expect(editor.snapshot().equation?.linear).toBe('A=πr^2');
  });

  test('a selected equation opens for editing; OK is one undo step', async () => {
    const editor = await open();
    const shape = editor.snapshot().shapes!.find((s) => s.name === 'Equation 3')!;
    editor.setSelection({ kind: 'shapes', nodeIds: [shape.nodeId] });
    const before = editor.snapshot().equation!.linear!;
    await show(editor);
    expect(title()).toBe('Edit equation');
    expect(field().value).toBe(before);
    await type('y=(a+1)/b');
    apply().click();
    await nextTick();
    expect(editor.snapshot().equation?.linear).toBe('y=(a+1)/b');
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().equation?.linear).toBe(before);
  });

  test('a linear form the editor cannot read blocks OK', async () => {
    const editor = await open();
    const state = await show(editor);
    await type('x^');
    apply().click();
    await nextTick();
    expect(state.closed).toBe(0);
    expect(document.querySelector('[role="alert"]')!.textContent).toContain('cannot read');
    expect(field().getAttribute('aria-invalid')).toBe('true');
  });
});
