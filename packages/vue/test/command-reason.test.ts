// `useEditorCommand` and `useEditorValueCommand` return the disabled reason as readable text in
// the root locale, as in React: a catalog key is translated, and the engine's own text shows as
// itself.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  createApp,
  defineComponent,
  h,
  nextTick,
  provide,
  shallowRef,
  type App,
  type Component,
  type ShallowRef,
} from 'vue';
import {
  chromeCommandState,
  createPptxEditor,
  setNotesTarget,
  translateKeyOrText,
  type ChromeActionHost,
  type NotesTarget,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { createTranslator, type LocaleCatalog, type Translate } from '@pptx-editor.dev/i18n';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import {
  PptxEditorRoot,
  PptxEditorViewport,
  pptxTranslationKey,
  useEditorCommand,
  useEditorValueCommand,
} from '../src/index.ts';
import { chromeActionHostKey, providePptxEditor } from '../src/composables.ts';

/** The English text of a catalog key. */
const en = createTranslator();

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
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

/** Mount `setup` under `editor`, a translator ref, and an optional action host. */
function mountWith(
  editor: PptxEditorInstance,
  setup: () => () => unknown,
  actions?: ChromeActionHost
): ShallowRef<Translate> {
  host = document.createElement('div');
  document.body.append(host);
  const translator = shallowRef<Translate>(createTranslator());
  const Child = defineComponent({ setup });
  const Provider = defineComponent({
    setup() {
      providePptxEditor(shallowRef(editor));
      provide(pptxTranslationKey, translator);
      if (actions) provide(chromeActionHostKey, actions);
      return () => h(Child);
    },
  });
  app = createApp(Provider as Component);
  app.mount(host);
  return translator;
}

describe('the disabled reason of useEditorCommand', () => {
  test('a missing root shows as text, and follows a locale change', async () => {
    const editor = await open();
    let cmd: ReturnType<typeof useEditorCommand> | null = null;
    const translator = mountWith(editor, () => {
      cmd = useEditorCommand('file.pageSetup');
      return () => h('span');
    });
    expect(cmd!.disabledReason.value).toBe(en('disabled.noRoot'));
    translator.value = createTranslator({ disabled: { noRoot: 'Kein Editor-Root' } });
    await nextTick();
    expect(cmd!.disabledReason.value).toBe('Kein Editor-Root');
  });

  test('a popup that the root turns off shows the reason in the root locale', async () => {
    let cmd: ReturnType<typeof useEditorCommand> | null = null;
    const Probe = defineComponent({
      setup() {
        cmd = useEditorCommand('find.find');
        return () => h('span');
      },
    });
    const locale = shallowRef<LocaleCatalog | undefined>(undefined);
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { popups: { find: false }, locale: locale.value }, () =>
            h(PptxEditorViewport, null, () => h(Probe))
          ),
      }) as Component
    );
    app.mount(host);
    await nextTick();
    expect(cmd!.disabledReason.value).toBe(en('find.unavailable'));
    locale.value = { find: { unavailable: 'Suchen ist aus' } };
    await nextTick();
    expect(cmd!.disabledReason.value).toBe('Suchen ist aus');
  });

  test("the engine's own text reason shows unchanged", async () => {
    const editor = await open();
    let cmd: ReturnType<typeof useEditorCommand> | null = null;
    mountWith(
      editor,
      () => {
        cmd = useEditorCommand('file.pageSetup');
        return () => h('span');
      },
      {
        openDialog: () => {},
        startSlideshow: () => {},
        openFind: () => {},
        openLink: () => {},
        pickImage: () => {},
        unavailable: () => 'The host is saving the file',
      }
    );
    expect(cmd!.disabledReason.value).toBe('The host is saving the file');
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
    let cmd: ReturnType<typeof useEditorCommand> | null = null;
    const translator = mountWith(editor, () => {
      cmd = useEditorCommand('text.bold');
      return () => h('span');
    });
    try {
      setNotesTarget(editor, target({ canFormat: false }));
      await nextTick();
      expect(cmd!.disabledReason.value).toBe(en('disabled.selectNotesText'));
      setNotesTarget(editor, target({ readOnly: true }));
      await nextTick();
      expect(cmd!.disabledReason.value).toBe(en('disabled.viewing'));
      translator.value = createTranslator({ disabled: { viewing: 'Nur Ansicht' } });
      await nextTick();
      expect(cmd!.disabledReason.value).toBe('Nur Ansicht');
      setNotesTarget(editor, target({}));
      await nextTick();
      expect(cmd!.disabledReason.value).toBe(null);
    } finally {
      setNotesTarget(editor, null);
    }
  });
});

describe('the disabled reason of useEditorValueCommand', () => {
  test('forwards the translated reason', async () => {
    const editor = await open();
    const key = chromeCommandState(editor, 'font.size').disabledReason;
    expect(key).toBeString();
    let cmd: ReturnType<typeof useEditorValueCommand<'font.size'>> | null = null;
    mountWith(editor, () => {
      cmd = useEditorValueCommand('font.size');
      return () => h('span');
    });
    expect(cmd!.disabledReason.value).toBe(translateKeyOrText(en, key!));
    expect(cmd!.disabledReason.value).not.toBe(key);
  });
});
