// `setFonts`: the scenarios of DOCX Editor's `packages/core/src/editor/__tests__/set-fonts.test.ts`
// (commit 5993bd223), on the PPTX session. Fonts change measurement, never content.

import { expect, test } from 'bun:test';
import {
  createPptxEditor,
  type EditorFontError,
  type FontConfigurationFragment,
  type FontResolutionRequest,
} from '../index.ts';
import { deferred, face, settle, missingFontsDeck } from './font-test-helpers.ts';

const substituted = (editor: ReturnType<typeof createPptxEditor>): string[] =>
  editor.snapshot().substitutedFonts.map((f) => f.family);

test('setFonts replaces host faces and keeps the document, selection, and undo history', async () => {
  const editor = createPptxEditor({ fonts: { sources: [face('Segoe UI')] } });
  await editor.load(missingFontsDeck());
  expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
  const session = editor.session;
  const selection = editor.snapshot().selection;
  const before = editor.session!.layout(0);
  expect(substituted(editor)).not.toContain('Segoe UI');

  await editor.setFonts({ sources: [face('Aptos')] });
  expect(editor.session).toBe(session);
  expect(editor.snapshot().selection).toEqual(selection);
  expect(editor.session!.layout(0)).not.toBe(before);
  expect(substituted(editor)).toContain('Segoe UI');
  expect(substituted(editor)).not.toContain('Aptos');
  expect(editor.snapshot().slides).toHaveLength(2);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(editor.snapshot().slides).toHaveLength(1);
  editor.destroy();
});

test('setFonts(undefined) removes the host faces and the substitutes return', async () => {
  const editor = createPptxEditor({ fonts: { sources: [face('Segoe UI')] } });
  await editor.load(missingFontsDeck());
  expect(substituted(editor)).not.toContain('Segoe UI');
  await editor.setFonts(undefined);
  expect(editor.session!.hostFaces.faces).toHaveLength(0);
  expect(substituted(editor)).toContain('Segoe UI');
  editor.destroy();
});

test('a replaced resolver is aborted and its late face cannot revive the retired font set', async () => {
  const gate = deferred<void>();
  let first: FontResolutionRequest | undefined;
  const editor = createPptxEditor();
  await editor.load(missingFontsDeck());
  const pending = editor.setFonts(async (request) => {
    first = request;
    await gate.promise;
    return { sources: [face('Segoe UI')] };
  });
  await settle();
  expect(first?.signal?.aborted).toBe(false);
  await editor.setFonts({ sources: [face('Aptos')] });
  expect(first?.signal?.aborted).toBe(true);
  gate.resolve();
  await pending;
  await settle();
  expect(substituted(editor)).toContain('Segoe UI');
  expect(substituted(editor)).not.toContain('Aptos');
  editor.destroy();
});

test('a replacement keeps the current faces until the new ones install', async () => {
  const gate = deferred<void>();
  const editor = createPptxEditor({ fonts: { sources: [face('Segoe UI')] } });
  await editor.load(missingFontsDeck());
  const pending = editor.setFonts(async () => {
    await gate.promise;
    return { sources: [face('Aptos')] };
  });
  await settle();
  expect(substituted(editor)).not.toContain('Segoe UI');
  gate.resolve();
  await pending;
  expect(substituted(editor)).toContain('Segoe UI');
  expect(substituted(editor)).not.toContain('Aptos');
  editor.destroy();
});

test('a replacement that supplies nothing usable still removes the replaced faces', async () => {
  const editor = createPptxEditor({ fonts: { sources: [face('Segoe UI')] } });
  await editor.load(missingFontsDeck());
  await editor.setFonts(() => undefined);
  expect(substituted(editor)).toContain('Segoe UI');
  editor.destroy();
});

test('a failing resolver removes the previous host faces and reports its cause', async () => {
  const errors: EditorFontError[] = [];
  const cause = new Error('font host down');
  const editor = createPptxEditor({
    fonts: { sources: [face('Segoe UI')] },
    onFontError: (e) => errors.push(e),
  });
  await editor.load(missingFontsDeck());
  await editor.setFonts(() => Promise.reject(cause));
  expect(substituted(editor)).toContain('Segoe UI');
  expect(errors).toHaveLength(1);
  expect(errors[0]!.cause).toBe(cause);
  editor.destroy();
});

test('loading a new document and destroying an editor abort their outstanding resolver', async () => {
  const signals: AbortSignal[] = [];
  const editor = createPptxEditor({
    fonts: (request): Promise<FontConfigurationFragment> => {
      if (request.signal) signals.push(request.signal);
      return new Promise(() => undefined);
    },
  });
  void editor.load(missingFontsDeck());
  await settle();
  void editor.load(missingFontsDeck());
  await settle();
  expect(signals).toHaveLength(2);
  expect(signals[0]!.aborted).toBe(true);
  editor.destroy();
  expect(signals[1]!.aborted).toBe(true);
});

test('setFonts after destroy throws', async () => {
  const editor = createPptxEditor();
  editor.destroy();
  await expect(editor.setFonts(undefined)).rejects.toThrow();
});

test('setFonts before a deck opens applies to the next deck', async () => {
  const editor = createPptxEditor();
  await editor.setFonts({ sources: [face('Segoe UI')] });
  await editor.load(missingFontsDeck());
  expect(substituted(editor)).not.toContain('Segoe UI');
  editor.destroy();
});
