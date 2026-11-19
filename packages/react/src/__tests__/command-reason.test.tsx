// `useEditorCommand` and `useEditorValueCommand` return the disabled reason as readable text in
// the root locale: a catalog key is translated, and the engine's own text shows as itself.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import {
  chromeCommandState,
  createPptxEditor,
  setNotesTarget,
  translateKeyOrText,
  type ChromeActionHost,
  type NotesTarget,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { createTranslator, type LocaleCatalog } from '@pptx-editor.dev/i18n';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import {
  ChromeActionHostContext,
  PptxEditorContext,
  TranslationContext,
  useEditorCommand,
  useEditorValueCommand,
} from '../context.ts';
import { PptxEditorRoot } from '../PptxEditorRoot.tsx';

/** The English text of a catalog key. */
const en = createTranslator();

afterEach(cleanup);

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

/** Render a probe for `slot` and return a reader of its last disabled reason. */
function probe(
  editor: PptxEditorInstance,
  slot: Parameters<typeof useEditorCommand>[0],
  wrap: (child: React.ReactNode) => React.ReactNode = (c) => c
) {
  const seen: { reason: string | null } = { reason: null };
  function Probe() {
    seen.reason = useEditorCommand(slot).disabledReason;
    return null;
  }
  const view = (locale?: LocaleCatalog) => (
    <TranslationContext.Provider value={createTranslator(locale)}>
      <PptxEditorContext.Provider value={editor}>{wrap(<Probe />)}</PptxEditorContext.Provider>
    </TranslationContext.Provider>
  );
  const r = render(view());
  return { seen, relocale: (locale?: LocaleCatalog) => r.rerender(view(locale)) };
}

describe('the disabled reason of useEditorCommand', () => {
  test('a missing root shows as text, and follows a locale change', async () => {
    const editor = await open();
    const { seen, relocale } = probe(editor, 'file.pageSetup');
    expect(seen.reason).toBe(en('disabled.noRoot'));
    relocale({ disabled: { noRoot: 'Kein Editor-Root' } });
    expect(seen.reason).toBe('Kein Editor-Root');
    relocale();
    expect(seen.reason).toBe(en('disabled.noRoot'));
  });

  test('a popup that the root turns off shows the reason in the root locale', async () => {
    const editor = await open();
    const seen: { reason: string | null } = { reason: null };
    function Probe() {
      seen.reason = useEditorCommand('find.find').disabledReason;
      return null;
    }
    const view = (locale?: LocaleCatalog) => (
      <PptxEditorRoot popups={{ find: false }} {...(locale ? { locale } : {})}>
        <PptxEditorContext.Provider value={editor}>
          <Probe />
        </PptxEditorContext.Provider>
      </PptxEditorRoot>
    );
    const r = render(view());
    expect(seen.reason).toBe(en('find.unavailable'));
    r.rerender(view({ find: { unavailable: 'Suchen ist aus' } }));
    expect(seen.reason).toBe('Suchen ist aus');
  });

  test("the engine's own text reason shows unchanged", async () => {
    const editor = await open();
    const host: ChromeActionHost = {
      openDialog: () => {},
      startSlideshow: () => {},
      openFind: () => {},
      openLink: () => {},
      pickImage: () => {},
      unavailable: () => 'The host is saving the file',
    };
    const { seen } = probe(editor, 'file.pageSetup', (c) => (
      <ChromeActionHostContext.Provider value={host}>{c}</ChromeActionHostContext.Provider>
    ));
    expect(seen.reason).toBe('The host is saving the file');
  });

  test('the notes formatting reasons show as text', async () => {
    const editor = await open();
    const target = (over: Partial<NotesTarget>): NotesTarget => ({
      formats: { bold: false, italic: false, underline: false, strike: false },
      readOnly: false,
      canFormat: true,
      apply: () => true,
      ...over,
    });
    const { seen, relocale } = probe(editor, 'text.bold');
    try {
      act(() => setNotesTarget(editor, target({ canFormat: false })));
      expect(seen.reason).toBe(en('disabled.selectNotesText'));
      act(() => setNotesTarget(editor, target({ readOnly: true })));
      expect(seen.reason).toBe(en('disabled.viewing'));
      relocale({ disabled: { viewing: 'Nur Ansicht' } });
      expect(seen.reason).toBe('Nur Ansicht');
      act(() => setNotesTarget(editor, target({})));
      expect(seen.reason).toBe(null);
    } finally {
      act(() => setNotesTarget(editor, null));
    }
  });
});

describe('the disabled reason of useEditorValueCommand', () => {
  test('forwards the translated reason', async () => {
    const editor = await open();
    const key = chromeCommandState(editor, 'font.size').disabledReason;
    expect(key).toBeString();
    const seen: { reason: string | null } = { reason: null };
    function Probe() {
      seen.reason = useEditorValueCommand('font.size').disabledReason;
      return null;
    }
    render(
      <PptxEditorContext.Provider value={editor}>
        <Probe />
      </PptxEditorContext.Provider>
    );
    expect(seen.reason).toBe(translateKeyOrText(en, key!));
    expect(seen.reason).not.toBe(key);
  });
});
