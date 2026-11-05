// The imperative `PptxEditorRef` that `PptxEditor` and `PptxEditorRoot` expose to a template
// ref, the twin of packages/react/src/__tests__/editor-ref.test.tsx and of DOCX Editor's Vue
// `DocxEditorRef`.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, h, nextTick, shallowRef, type App, type Component } from 'vue';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditor, PptxEditorContent, PptxEditorRoot, type PptxEditorRef } from '../src/index.ts';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

/** Mount `component` with a template ref; the ref holds the exposed handle. */
function mount(component: Component, props: Record<string, unknown>, children?: () => unknown) {
  const handle = shallowRef<PptxEditorRef | null>(null);
  host = document.createElement('div');
  document.body.append(host);
  app = createApp({
    setup: () => () =>
      h(component, { ...props, ref: handle }, children ? { default: children } : undefined),
  });
  app.mount(host);
  return handle;
}

async function ready(handle: { value: PptxEditorRef | null }): Promise<void> {
  for (let i = 0; i < 200 && handle.value?.snapshot().status !== 'ready'; i += 1)
    await new Promise((r) => setTimeout(r, 10));
  await nextTick();
}

describe('PptxEditorRef', () => {
  test('PptxEditorRoot exposes the ref members over its instance', async () => {
    const handle = mount(PptxEditorRoot, { document: buildSlideLayoutsDeck() }, () =>
      h(PptxEditorContent)
    );
    await ready(handle);
    const ref = handle.value!;
    const editor = ref.getEditor()!;
    expect(ref.snapshot()).toBe(editor.snapshot());
    const bytes = ref.save();
    expect(bytes instanceof Uint8Array && bytes.length > 0).toBe(true);
    ref.focus();
    expect((await ref.load(bytes!)).ok).toBe(true);
    expect(ref.getEditor() === editor).toBe(true);
    app!.unmount();
    app = null;
    // After unmount the members answer without an editor, as before mount.
    expect(ref.getEditor() === null).toBe(true);
    expect(ref.exec({ type: 'undo' })).toEqual({ ok: false, reason: 'disabled.noEditor' });
    expect(ref.snapshot().status).toBe('empty');
  });

  test('PptxEditor exposes the same members through its root', async () => {
    const handle = mount(PptxEditor, { document: buildSlideLayoutsDeck() });
    await ready(handle);
    const ref = handle.value!;
    expect(ref.snapshot().slides.length > 0).toBe(true);
    const before = ref.snapshot().revision;
    expect(ref.exec({ type: 'insertTextBox' }).ok).toBe(true);
    expect(ref.snapshot().revision > before).toBe(true);
    expect(Object.keys(ref).sort()).toEqual(
      ['exec', 'focus', 'getEditor', 'load', 'save', 'snapshot'].sort()
    );
  });
});
