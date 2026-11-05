// Late faces (`FontConfigurationFragment.deferred`, `deferFonts`): the scenarios of DOCX
// Editor's `packages/core/src/editor/__tests__/deferred-fonts.test.ts` (commit 5993bd223), on
// the PPTX session.

import { describe, expect, test } from 'bun:test';
import {
  createPptxEditor,
  deferFonts,
  defineFontResolver,
  type EditorFontError,
  type FontConfigurationFragment,
} from '../index.ts';
import { deferred, face, fixture, settle } from './font-test-helpers.ts';

const substituted = (editor: ReturnType<typeof createPptxEditor>): string[] =>
  editor.snapshot().substitutedFonts.map((f) => f.family);

describe('deferred font faces in the editor', () => {
  test('the immediate faces install first, the late ones on arrival', async () => {
    const late = deferred<FontConfigurationFragment | undefined>();
    const editor = createPptxEditor({
      fonts: { sources: [face('Segoe UI')], deferred: late.promise },
    });
    await editor.load(fixture('font-substitution.pptx'));
    expect(substituted(editor)).not.toContain('Segoe UI');
    expect(substituted(editor)).toContain('Aptos');
    const generation = editor.session!.fontGeneration;
    late.resolve({ sources: [face('Aptos')] });
    await settle();
    expect(editor.session!.fontGeneration).toBeGreaterThan(generation);
    expect(substituted(editor)).not.toContain('Aptos');
    expect(substituted(editor)).not.toContain('Segoe UI');
    editor.destroy();
  });

  test('a late answer for a replaced document never installs', async () => {
    const late = deferred<FontConfigurationFragment | undefined>();
    let calls = 0;
    const editor = createPptxEditor({
      fonts: defineFontResolver(() =>
        ++calls === 1 ? { sources: [], deferred: late.promise } : undefined
      ),
    });
    await editor.load(fixture('font-substitution.pptx'));
    await editor.load(fixture('font-substitution.pptx'));
    late.resolve({ sources: [face('Aptos')] });
    await settle();
    expect(substituted(editor)).toContain('Aptos');
    editor.destroy();
  });

  test('a late promise that rejects ends the wait', async () => {
    const editor = createPptxEditor({
      fonts: { sources: [face('Segoe UI')], deferred: Promise.reject(new Error('late failure')) },
      onFontError: () => undefined,
    });
    expect((await editor.load(fixture('font-substitution.pptx'))).ok).toBe(true);
    await settle();
    expect(substituted(editor)).not.toContain('Segoe UI');
    editor.destroy();
  });

  test('deferFonts as the only resolver installs its faces late and reports its failures', async () => {
    const errors: EditorFontError[] = [];
    const editor = createPptxEditor({
      fonts: deferFonts(
        defineFontResolver(async () => ({
          sources: [face('Aptos')],
          failures: [new Error('one face missing')],
        }))
      ),
      onFontError: (e) => errors.push(e),
    });
    await editor.load(fixture('font-substitution.pptx'));
    await settle(10);
    expect(substituted(editor)).not.toContain('Aptos');
    expect(errors.map((e) => e.diagnostic)).toContain('one face missing');
    editor.destroy();
  });

  test('setFonts with only background faces keeps waiting for them', async () => {
    const late = deferred<FontConfigurationFragment | undefined>();
    const editor = createPptxEditor();
    await editor.load(fixture('font-substitution.pptx'));
    await editor.setFonts({ sources: [], deferred: late.promise });
    expect(substituted(editor)).toContain('Segoe UI');
    late.resolve({ sources: [face('Segoe UI')] });
    await settle();
    expect(substituted(editor)).not.toContain('Segoe UI');
    editor.destroy();
  });
});
