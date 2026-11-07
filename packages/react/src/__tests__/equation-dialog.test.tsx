import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import {
  createPptxEditor,
  toolbarCommandState,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { PptxEditorContext } from '../context.ts';
import { EquationDialog } from '../index.ts';

// The Equation dialog: a gallery pick inserts an equation, the linear form of a selected
// equation edits it as one undo step, and a linear form the editor cannot read blocks OK.

const FIXTURE = join(import.meta.dir, '../../../../fixtures/equations.pptx');

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(new Uint8Array(readFileSync(FIXTURE)))).ok).toBe(true);
  act(() => editor.selectSlide(1));
  return editor;
}

function show(editor: PptxEditorInstance, onClose: () => void = () => {}) {
  return render(
    <PptxEditorContext.Provider value={editor}>
      <EquationDialog open onClose={onClose} />
    </PptxEditorContext.Provider>
  );
}

const field = (): HTMLInputElement => screen.getByLabelText('Linear form') as HTMLInputElement;
const apply = (): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>('[data-pptx-part="apply"]')!;

afterEach(cleanup);

describe('Equation dialog', () => {
  test('a gallery pick fills the linear form, and Insert adds the equation', async () => {
    const editor = await open();
    expect(toolbarCommandState(editor, 'insert.equation')).toMatchObject({ enabled: true });
    let closed = 0;
    show(editor, () => (closed += 1));
    expect(screen.getByRole('heading', { name: 'Insert equation' })).toBeDefined();
    fireEvent.click(screen.getByText('Area of a circle'));
    expect(field().value).toBe('A=\\pi r^2');
    act(() => apply().click());
    expect(closed).toBe(1);
    expect(editor.snapshot().equation?.linear).toBe('A=πr^2');
  });

  test('a selected equation opens for editing; OK is one undo step', async () => {
    const editor = await open();
    const shape = editor.snapshot().shapes!.find((s) => s.name === 'Equation 3')!;
    act(() => editor.setSelection({ kind: 'shapes', nodeIds: [shape.nodeId] }));
    const before = editor.snapshot().equation!.linear!;
    show(editor);
    expect(screen.getByRole('heading', { name: 'Edit equation' })).toBeDefined();
    expect(field().value).toBe(before);
    fireEvent.change(field(), { target: { value: 'y=(a+1)/b' } });
    act(() => apply().click());
    expect(editor.snapshot().equation?.linear).toBe('y=(a+1)/b');
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().equation?.linear).toBe(before);
  });

  test('a linear form the editor cannot read blocks OK', async () => {
    const editor = await open();
    let closed = 0;
    show(editor, () => (closed += 1));
    fireEvent.change(field(), { target: { value: 'x^' } });
    act(() => apply().click());
    expect(closed).toBe(0);
    expect(screen.getByRole('alert').textContent).toContain('cannot read');
    expect(field().getAttribute('aria-invalid')).toBe('true');
  });
});
