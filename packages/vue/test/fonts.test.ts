// The Vue twin of the `fonts` prop and `useFonts` (packages/react/src/__tests__/fonts.test.tsx).

import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { face, missingFontsDeck } from '../../core/src/editor/__tests__/font-test-helpers.ts';
import {
  PptxEditor,
  PptxEditorRoot,
  defineFontResolver,
  isFontResolver,
  useFonts,
  type EditorFontError,
  type FontConfigurationFragment,
  type FontResolutionRequest,
} from '../src/index.ts';

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

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

const deck = missingFontsDeck();
const substituted = (e: PptxEditorInstance) => e.snapshot().substitutedFonts.map((f) => f.family);

function mount(component: Component): void {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(component);
  app.mount(host);
}

async function ready(ref: { editor: PptxEditorInstance | null }): Promise<PptxEditorInstance> {
  for (let i = 0; i < 300 && ref.editor?.snapshot().status !== 'ready'; i += 1)
    await new Promise((r) => setTimeout(r, 10));
  return ref.editor!;
}

describe('the fonts prop', () => {
  test('a new fonts value replaces the fonts in place and keeps the document', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const fonts = shallowRef<FontConfigurationFragment>({ sources: [face('Segoe UI')] });
    mount(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, {
            document: deck,
            fonts: fonts.value,
            onReady: (e: PptxEditorInstance) => (ref.editor = e),
          }),
      })
    );
    const editor = await ready(ref);
    expect(substituted(editor)).not.toContain('Segoe UI');
    const session = editor.session;
    fonts.value = { sources: [face('Aptos')] };
    await nextTick();
    await new Promise((r) => setTimeout(r, 30));
    expect(editor.session).toBe(session);
    expect(substituted(editor)).toContain('Segoe UI');
    expect(substituted(editor)).not.toContain('Aptos');
  });

  test('a failing origin emits fontError from the sugar component', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const errors: EditorFontError[] = [];
    mount(
      defineComponent({
        setup: () => () =>
          h(PptxEditor, {
            document: deck,
            fonts: () => Promise.reject(new Error('offline')),
            onFontError: (e: EditorFontError) => errors.push(e),
            onReady: (e: PptxEditorInstance) => (ref.editor = e),
          }),
      })
    );
    await ready(ref);
    expect(errors.map((e) => e.diagnostic)).toEqual(['offline']);
  });
});

describe('useFonts', () => {
  test('returns a marked resolver that composes its origins first-wins', async () => {
    let resolver: ReturnType<typeof useFonts> | null = null;
    mount(
      defineComponent({
        setup() {
          resolver = useFonts(
            { sources: [face('Segoe UI')] },
            defineFontResolver((request: FontResolutionRequest) => ({
              sources: request.families.includes('Aptos') ? [face('Aptos')] : [],
            }))
          );
          return () => null;
        },
      })
    );
    expect(isFontResolver(resolver)).toBe(true);
    const answer = await resolver!({ families: ['Aptos', 'Segoe UI'], defaultFamily: 'Calibri' });
    expect(answer?.sources?.map((s) => s.request.family).sort()).toEqual(['Aptos', 'Segoe UI']);
  });
});
