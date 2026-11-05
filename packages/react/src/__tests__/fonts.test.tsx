// The `fonts` prop and `useFonts`, after DOCX Editor's root: a stable `useFonts` value never
// replaces the fonts, a new `fonts` value replaces them in place (`setFonts`), and failures
// reach `onFontError`.

import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { useState } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import type {
  EditorFontError,
  FontResolutionRequest,
  PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { face, fixture } from '../../../core/src/editor/__tests__/font-test-helpers.ts';
import { PptxEditor, useFonts, type FontConfigurationFragment } from '../index.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';

afterEach(cleanup);

// happy-dom has no FontFace API: host faces register with the page, so give it a minimal one.
const scope = globalThis as unknown as { FontFace?: unknown };
const hadFontFace = scope.FontFace !== undefined;
const fontSet = (document as unknown as { fonts?: unknown }).fonts;
beforeAll(() => {
  if (!hadFontFace)
    scope.FontFace = class {
      load = () => Promise.resolve(this);
    };
  if (!fontSet)
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: { add: () => undefined, delete: () => true },
    });
});
afterAll(() => {
  if (!hadFontFace) delete scope.FontFace;
  if (!fontSet) Reflect.deleteProperty(document, 'fonts');
});

const deck = fixture('font-substitution.pptx');

async function ready(ref: { editor: PptxEditorInstance | null }): Promise<PptxEditorInstance> {
  for (let i = 0; i < 300 && ref.editor?.snapshot().status !== 'ready'; i += 1)
    await act(() => new Promise((r) => setTimeout(r, 10)));
  return ref.editor!;
}

const wait = () => act(() => new Promise((r) => setTimeout(r, 20)));
const substituted = (e: PptxEditorInstance) => e.snapshot().substitutedFonts.map((f) => f.family);

describe('the fonts prop', () => {
  test('a host face from the fonts prop covers its family', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    render(
      <PptxEditorRoot
        document={deck}
        fonts={{ sources: [face('Segoe UI')] }}
        onReady={(e) => (ref.editor = e)}
      >
        <PptxEditorContent />
      </PptxEditorRoot>
    );
    const editor = await ready(ref);
    expect(substituted(editor)).not.toContain('Segoe UI');
  });

  test('a new fonts value replaces the fonts in place and keeps the document', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const view = (fonts: FontConfigurationFragment) => (
      <PptxEditorRoot document={deck} fonts={fonts} onReady={(e) => (ref.editor = e)}>
        <PptxEditorContent />
      </PptxEditorRoot>
    );
    const { rerender } = render(view({ sources: [face('Segoe UI')] }));
    const editor = await ready(ref);
    const session = editor.session;
    rerender(view({ sources: [face('Aptos')] }));
    await wait();
    expect(ref.editor).toBe(editor);
    expect(editor.session).toBe(session);
    expect(substituted(editor)).toContain('Segoe UI');
    expect(substituted(editor)).not.toContain('Aptos');
  });

  test('useFonts keeps one identity across renders, so the fonts are not replaced', async () => {
    const ref: { editor: PptxEditorInstance | null; rerender: (() => void) | null } = {
      editor: null,
      rerender: null,
    };
    const requests: FontResolutionRequest[] = [];
    function Host() {
      const [, setTick] = useState(0);
      ref.rerender = () => setTick((n) => n + 1);
      // An inline fragment each render: useFonts still returns one resolver.
      const fonts = useFonts(
        (request: FontResolutionRequest) => (
          requests.push(request),
          { sources: [face('Segoe UI')] }
        )
      );
      return (
        <PptxEditorRoot document={deck} fonts={fonts} onReady={(e) => (ref.editor = e)}>
          <PptxEditorContent />
        </PptxEditorRoot>
      );
    }
    render(<Host />);
    const editor = await ready(ref);
    expect(requests).toHaveLength(1);
    act(() => ref.rerender!());
    act(() => ref.rerender!());
    await wait();
    expect(requests).toHaveLength(1);
    expect(substituted(editor)).not.toContain('Segoe UI');
  });

  test('a failing origin reaches onFontError on the sugar component', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const errors: EditorFontError[] = [];
    render(
      <PptxEditor
        document={deck}
        fonts={() => Promise.reject(new Error('offline'))}
        onFontError={(e) => errors.push(e)}
        onReady={(e) => (ref.editor = e)}
      />
    );
    await ready(ref);
    expect(errors.map((e) => e.diagnostic)).toEqual(['offline']);
  });
});
