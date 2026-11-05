// The imperative `PptxEditorRef` of `PptxEditor` and `PptxEditor.Root`, as DOCX Editor's
// `DocxEditorRef`: every member forwards to the instance, and is safe before mount and after
// unmount. The Vue twin is packages/vue/test/editor-ref.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { createRef } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createEditorRefApi } from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditor, type PptxEditorRef } from '../index.ts';

afterEach(cleanup);

async function ready(ref: { current: PptxEditorRef | null }): Promise<void> {
  for (let i = 0; i < 200 && ref.current?.snapshot().status !== 'ready'; i += 1)
    await act(() => new Promise((r) => setTimeout(r, 10)));
}

describe('PptxEditorRef', () => {
  test('before mount, every member answers without an editor', async () => {
    const api = createEditorRefApi(() => null);
    expect(api.getEditor() === null).toBe(true);
    expect(api.save() === null).toBe(true);
    expect(api.exec({ type: 'undo' })).toEqual({ ok: false, reason: 'disabled.noEditor' });
    expect(await api.load(new Uint8Array())).toEqual({ ok: false, reason: 'disabled.noEditor' });
    expect(api.snapshot().status).toBe('empty');
    api.focus();
  });

  test('PptxEditor.Root forwards the ref to its instance', async () => {
    const ref = createRef<PptxEditorRef>();
    const view = render(
      <PptxEditor.Root ref={ref} document={buildSlideLayoutsDeck()}>
        <PptxEditor.Content />
      </PptxEditor.Root>
    );
    await ready(ref);
    const editor = ref.current!.getEditor()!;
    expect(ref.current!.snapshot()).toBe(editor.snapshot());
    const bytes = ref.current!.save();
    expect(bytes instanceof Uint8Array && bytes.length > 0).toBe(true);
    act(() => ref.current!.focus());
    // `load` opens new bytes through the same instance.
    expect((await act(() => ref.current!.load(bytes!))).ok).toBe(true);
    expect(ref.current!.getEditor() === editor).toBe(true);
    view.unmount();
    expect(ref.current === null).toBe(true);
  });

  test('PptxEditor forwards the ref through its root', async () => {
    const ref = createRef<PptxEditorRef>();
    render(<PptxEditor ref={ref} document={buildSlideLayoutsDeck()} />);
    await ready(ref);
    expect(ref.current!.snapshot().slides.length > 0).toBe(true);
    const before = ref.current!.snapshot().revision;
    let result: { ok: boolean } = { ok: false };
    act(() => void (result = ref.current!.exec({ type: 'insertTextBox' })));
    expect(result.ok).toBe(true);
    expect(ref.current!.snapshot().revision > before).toBe(true);
    expect(ref.current!.snapshot().canUndo).toBe(true);
  });
});
