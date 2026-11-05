// The Vue twin of the `modules` and `fontFamilies` props
// (packages/react/src/__tests__/root-options.test.tsx): `modules` reaches `createPptxEditor` at
// mount, and a new `fontFamilies` list updates the font picker in place.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  resetDevWarnings,
  type CollaborationDocumentPort,
  type EditorCollaborationSession,
  type EditorModule,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditor, PptxEditorRoot } from '../src/index.ts';

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

const IN = 914_400;
const deck = buildPptx({
  slides: [
    {
      layout: 'blank',
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
    },
  ],
});

function mount(component: Component): void {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(component);
  app.mount(host);
}

/** A collaboration module whose session records the ports the editor attaches. */
function recordingModule(): { module: EditorModule; ports: CollaborationDocumentPort[] } {
  const ports: CollaborationDocumentPort[] = [];
  const session: EditorCollaborationSession = {
    documentId: 'doc',
    sessionId: 'session',
    identity: { actorId: 'a', name: 'A' },
    attached: false,
    status: () => 'ready',
    statusSnapshot: () => ({ status: 'ready', reason: undefined, lastFailure: undefined }),
    subscribeStatus: () => () => {},
    participants: () => [],
    subscribeParticipants: () => () => {},
    destroy: () => {},
    attach(port) {
      ports.push(port);
      return () => {};
    },
  };
  return { module: { id: 'collaboration', collaboration: { session } }, ports };
}

async function ready(ref: { editor: PptxEditorInstance | null }): Promise<PptxEditorInstance> {
  for (let i = 0; i < 300 && ref.editor?.snapshot().status !== 'ready'; i += 1)
    await new Promise((r) => setTimeout(r, 10));
  return ref.editor!;
}

describe('the modules prop (Vue)', () => {
  test('PptxEditorRoot passes modules to the instance it creates', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const { module, ports } = recordingModule();
    mount(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, {
            document: deck,
            modules: [module],
            onReady: (e: PptxEditorInstance) => (ref.editor = e),
          }),
      })
    );
    const editor = await ready(ref);
    expect(ports.length).toBe(1);
    expect(ports[0]!.package()).toBe(editor.session!.store.package);
  });

  test('PptxEditor forwards modules to its root', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const { module, ports } = recordingModule();
    mount(
      defineComponent({
        setup: () => () =>
          h(PptxEditor, {
            document: deck,
            modules: [module],
            onReady: (e: PptxEditorInstance) => (ref.editor = e),
          }),
      })
    );
    await ready(ref);
    expect(ports.length).toBe(1);
  });

  test('a modules change after mount warns once; a new array of the same modules does not', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const ref: { editor: PptxEditorInstance | null } = { editor: null };
      const { module, ports } = recordingModule();
      const modules = shallowRef<readonly EditorModule[]>([module]);
      mount(
        defineComponent({
          setup: () => () =>
            h(PptxEditorRoot, {
              document: deck,
              modules: modules.value,
              onReady: (e: PptxEditorInstance) => (ref.editor = e),
            }),
        })
      );
      const editor = await ready(ref);
      const modulesWarnings = () =>
        warn.mock.calls.filter((c) => String(c[0]).includes('modules change')).length;
      // A new array that holds the same module objects is no change.
      modules.value = [module];
      await nextTick();
      await nextTick();
      expect(modulesWarnings()).toBe(0);
      modules.value = [module, recordingModule().module];
      await nextTick();
      await nextTick();
      expect(modulesWarnings()).toBe(1);
      expect(String(warn.mock.calls.at(-1)?.[0])).toContain('PptxEditorRoot');
      // The instance and its modules stay as they were created.
      expect(ref.editor).toBe(editor);
      expect(ports.length).toBe(1);
    } finally {
      warn.mockRestore();
    }
  });
});

describe('the fontFamilies prop (Vue)', () => {
  test('a new list updates the font picker in place and keeps the document', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const families = shallowRef<readonly string[] | undefined>(['Inter']);
    mount(
      defineComponent({
        setup: () => () =>
          h(PptxEditor, {
            document: deck,
            ...(families.value ? { fontFamilies: families.value } : {}),
            onReady: (e: PptxEditorInstance) => (ref.editor = e),
          }),
      })
    );
    const editor = await ready(ref);
    const session = editor.session;
    expect(editor.fontFamilies()).toContain('Inter');
    families.value = ['Lato'];
    await nextTick();
    await nextTick();
    expect(ref.editor).toBe(editor);
    expect(editor.session).toBe(session);
    expect(editor.fontFamilies()).toContain('Lato');
    expect(editor.fontFamilies()).not.toContain('Inter');
    // A new array with the same families changes nothing.
    const before = editor.fontFamilies();
    families.value = ['Lato'];
    await nextTick();
    await nextTick();
    expect(editor.fontFamilies()).toBe(before);
    families.value = undefined;
    await nextTick();
    await nextTick();
    expect(editor.fontFamilies()).not.toContain('Lato');
  });
});
