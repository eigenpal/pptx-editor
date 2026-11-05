// The root's locale reaches the canvas: a locale change repaints the placeholder prompts once,
// without a reload.

import { afterEach, describe, expect, test } from 'bun:test';
import { useState } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';

afterEach(cleanup);

describe('the root locale', () => {
  test('a locale change repaints the canvas prompts', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const deck = buildSlideLayoutsDeck();
    const view = (locale?: Record<string, unknown>) => (
      <PptxEditorRoot
        document={deck}
        onReady={(e) => (ref.editor = e)}
        {...(locale ? { locale } : {})}
      >
        <PptxEditorContent />
      </PptxEditorRoot>
    );
    const { rerender } = render(view());
    for (let i = 0; i < 200 && ref.editor?.snapshot().status !== 'ready'; i += 1)
      await act(() => new Promise((r) => setTimeout(r, 10)));
    // Slide 5 has an empty title: the canvas shows its prompt.
    act(() => ref.editor!.goToSlide(4));
    const canvas = (): string => document.querySelector('.pptx-surface__stage')?.textContent ?? '';
    expect(canvas()).toContain('Click to add title');
    rerender(view({ canvas: { prompts: { title: 'Titel hinzufügen' } } }));
    expect(canvas()).toContain('Titel hinzufügen');
    expect(canvas()).not.toContain('Click to add title');
  });

  test('an inline locale with an onChange that sets state renders without a loop', async () => {
    const ref: { editor: PptxEditorInstance | null; renders: number } = {
      editor: null,
      renders: 0,
    };
    const deck = buildSlideLayoutsDeck();
    const errors: string[] = [];
    const error = console.error;
    console.error = (...args: unknown[]) => void errors.push(args.map(String).join(' '));
    function Host() {
      const [, setChanges] = useState(0);
      ref.renders += 1;
      return (
        <PptxEditorRoot
          document={deck}
          // A new object with the same content on each render.
          locale={{ canvas: { prompts: { title: 'Titel hinzufügen' } } }}
          onReady={(e) => (ref.editor = e)}
          onChange={() => setChanges((n) => n + 1)}
        >
          <PptxEditorContent />
        </PptxEditorRoot>
      );
    }
    try {
      render(<Host />);
      for (let i = 0; i < 200 && ref.editor?.snapshot().status !== 'ready'; i += 1)
        await act(() => new Promise((r) => setTimeout(r, 10)));
      act(() => ref.editor!.goToSlide(4));
      await act(() => new Promise((r) => setTimeout(r, 50)));
      const settled = ref.renders;
      await act(() => new Promise((r) => setTimeout(r, 50)));
      expect(ref.renders).toBe(settled);
      expect(errors.some((e) => e.includes('Maximum update depth'))).toBe(false);
      expect(document.querySelector('.pptx-surface__stage')?.textContent).toContain(
        'Titel hinzufügen'
      );
    } finally {
      console.error = error;
    }
  });
});
