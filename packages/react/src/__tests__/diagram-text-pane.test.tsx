// The Diagram tab of the side panel in the React adapter: the empty state, the node fields that
// apply on Enter as one undo step each, add and delete, and the diagram context menu row and
// slot that turn the side panel to the tab.
// `packages/vue/test/diagram-text-pane.test.ts` checks that the Vue twin renders the same markup.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditorContent, PptxEditorRoot } from '../PptxEditorRoot.tsx';
import { SidePanel, Toolbar } from '../index.ts';

const deck = new Uint8Array(
  readFileSync(join(import.meta.dir, '../../../../fixtures/diagram-editing.pptx'))
);

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

async function mount(ui: React.ReactNode): Promise<PptxEditorInstance> {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)}>
      <PptxEditorContent />
      {ui}
    </PptxEditorRoot>
  );
  await act(async () => void (await editor!.load(deck)));
  return editor!;
}

function selectDiagram(editor: PptxEditorInstance): string {
  const frame = editor.snapshot().shapes!.find((s) => s.name.startsWith('Diagram'))!.nodeId;
  act(() => editor.setSelection({ kind: 'shapes', nodeIds: [frame] }));
  return frame;
}

const texts = (editor: PptxEditorInstance): string[] =>
  editor.snapshot().diagram!.nodes.map((n) => n.text);

describe('Diagram tab (React)', () => {
  test('the tab shows only for a diagram, and Edit text opens it', async () => {
    const editor = await mount(
      <>
        <Toolbar preset={false}>
          <Toolbar.Slot slotId="diagram.editText" />
        </Toolbar>
        <SidePanel />
      </>
    );
    expect(screen.queryByRole('tab', { name: 'Diagram' })).toBe(null);
    selectDiagram(editor);
    const tab = screen.getByRole('tab', { name: 'Diagram' });
    expect(tab.getAttribute('aria-selected')).toBe('false');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Edit text…' })));
    expect(tab.getAttribute('aria-selected')).toBe('true');
    expect(screen.getByRole('textbox', { name: 'Text of shape 1' })).toBeTruthy();
    // With the diagram deselected, the open tab stays and says to select a diagram.
    act(() => editor.setSelection({ kind: 'none' }));
    expect(screen.getByText('No diagram selected')).toBeTruthy();
  });

  test('a field applies on Enter as one undo step; Shift+Enter keeps typing', async () => {
    const editor = await mount(<SidePanel tab="diagram" />);
    selectDiagram(editor);
    const before = texts(editor);
    const field = screen.getByRole('textbox', { name: 'Text of shape 1' }) as HTMLTextAreaElement;
    expect(field.value).toBe(before[0]!);
    act(() => void fireEvent.change(field, { target: { value: 'Scope' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Enter', shiftKey: true }));
    expect(texts(editor)[0]).toBe(before[0]!);
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    expect(texts(editor)[0]).toBe('Scope');
    expect(field.value).toBe('Scope');
    act(() => void editor.exec({ type: 'undo' }));
    expect(texts(editor)[0]).toBe(before[0]!);
    // Escape drops a draft.
    act(() => void fireEvent.change(field, { target: { value: 'Draft' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Escape' }));
    expect(field.value).toBe(before[0]!);
  });

  test('adds after a node and deletes it', async () => {
    const editor = await mount(<SidePanel tab="diagram" />);
    selectDiagram(editor);
    const before = texts(editor);
    const first = before[0]!;
    act(
      () => void fireEvent.click(screen.getByRole('button', { name: `Add shape after ${first}` }))
    );
    expect(texts(editor)).toEqual([first, '', ...before.slice(1)]);
    // The focus goes to the new shape's field, so the user can type its text.
    expect(document.activeElement).toBe(screen.getByRole('textbox', { name: 'Text of shape 2' }));
    act(
      () => void fireEvent.click(screen.getByRole('button', { name: 'Delete shape Empty shape' }))
    );
    expect(texts(editor)).toEqual(before);
    // After a delete, the focus goes to the field that took the deleted row's place.
    expect(document.activeElement).toBe(screen.getByRole('textbox', { name: 'Text of shape 2' }));
    // The process has no children: Add shape below is off.
    const below = screen.getByRole('button', { name: `Add shape below ${first}` });
    expect(below.hasAttribute('disabled') || below.getAttribute('aria-disabled') === 'true').toBe(
      true
    );
  });
});
