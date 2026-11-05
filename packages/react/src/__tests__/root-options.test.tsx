// The `modules` and `fontFamilies` props of `PptxEditor` and `PptxEditor.Root`, after DOCX
// Editor's root: `modules` reaches `createPptxEditor` at mount (a collaboration module's
// session attaches when the deck opens), and a new `fontFamilies` list updates the font picker
// in place, keeping the instance and the document.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import {
  resetDevWarnings,
  type CollaborationDocumentPort,
  type EditorCollaborationSession,
  type EditorModule,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor } from '../index.ts';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';

afterEach(cleanup);

const IN = 914_400;
const deck = buildPptx({
  slides: [
    {
      layout: 'blank',
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
    },
  ],
});

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
    await act(() => new Promise((r) => setTimeout(r, 10)));
  return ref.editor!;
}

describe('the modules prop', () => {
  test('PptxEditor.Root passes modules to the instance it creates', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const { module, ports } = recordingModule();
    render(
      <PptxEditorRoot document={deck} modules={[module]} onReady={(e) => (ref.editor = e)}>
        <PptxEditorContent />
      </PptxEditorRoot>
    );
    const editor = await ready(ref);
    expect(ports.length).toBe(1);
    expect(ports[0]!.package()).toBe(editor.session!.store.package);
  });

  test('PptxEditor forwards modules to its root', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const { module, ports } = recordingModule();
    render(<PptxEditor document={deck} modules={[module]} onReady={(e) => (ref.editor = e)} />);
    await ready(ref);
    expect(ports.length).toBe(1);
  });

  test('a modules change after mount warns once; an inline array of the same modules does not', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const ref: { editor: PptxEditorInstance | null } = { editor: null };
      const { module, ports } = recordingModule();
      const view = (modules: readonly EditorModule[]) => (
        <PptxEditorRoot document={deck} modules={modules} onReady={(e) => (ref.editor = e)}>
          <PptxEditorContent />
        </PptxEditorRoot>
      );
      const { rerender } = render(view([module]));
      const editor = await ready(ref);
      // A new array that holds the same module objects is no change.
      rerender(view([module]));
      await act(async () => {});
      const modulesWarnings = () =>
        warn.mock.calls.filter((c) => String(c[0]).includes('modules change')).length;
      expect(modulesWarnings()).toBe(0);
      rerender(view([module, recordingModule().module]));
      await act(async () => {});
      expect(modulesWarnings()).toBe(1);
      expect(String(warn.mock.calls.at(-1)?.[0])).toContain('PptxEditor.Root');
      // The instance and its modules stay as they were created.
      expect(ref.editor).toBe(editor);
      expect(ports.length).toBe(1);
    } finally {
      warn.mockRestore();
    }
  });
});

describe('the fontFamilies prop', () => {
  test('a new list updates the font picker in place and keeps the document', async () => {
    const ref: { editor: PptxEditorInstance | null } = { editor: null };
    const view = (fontFamilies: readonly string[] | undefined) => (
      <PptxEditor
        document={deck}
        {...(fontFamilies ? { fontFamilies } : {})}
        onReady={(e) => (ref.editor = e)}
      />
    );
    const { rerender } = render(view(['Inter']));
    const editor = await ready(ref);
    const session = editor.session;
    expect(editor.fontFamilies()).toContain('Inter');
    rerender(view(['Lato']));
    await act(async () => {});
    expect(ref.editor).toBe(editor);
    expect(editor.session).toBe(session);
    expect(editor.fontFamilies()).toContain('Lato');
    expect(editor.fontFamilies()).not.toContain('Inter');
    // An inline array with the same families changes nothing.
    const before = editor.fontFamilies();
    rerender(view(['Lato']));
    await act(async () => {});
    expect(editor.fontFamilies()).toBe(before);
    rerender(view(undefined));
    await act(async () => {});
    expect(editor.fontFamilies()).not.toContain('Lato');
  });
});
