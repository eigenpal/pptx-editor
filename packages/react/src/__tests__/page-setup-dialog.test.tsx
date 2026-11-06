import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactNode } from 'react';
import { strFromU8, unzipSync } from 'fflate';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PageSetupDialog } from '../index.ts';

// The slide size fields of the Page setup dialog: a preset, typed inches, the orientation,
// the content choice, the range error, and one undo step. The Vue twin is
// packages/vue/test/page-setup-dialog.test.ts.

const BOX =
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1219200" y="685800"/><a:ext cx="2438400" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>';

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildPptx({ slides: [{ layout: 'blank', shapes: BOX }] }))).ok).toBe(
    true
  );
  return editor;
}

function withEditor(editor: PptxEditorInstance, ui: ReactNode) {
  return render(<PptxEditorContext.Provider value={editor}>{ui}</PptxEditorContext.Provider>);
}

const field = (name: string) => document.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
const ok = () => screen.getByRole('button', { name: 'OK' });
/** The saved width of the Box shape, in EMU. */
function boxWidth(editor: PptxEditorInstance): number {
  const slide = strFromU8(unzipSync(editor.save()!)['ppt/slides/slide1.xml']!);
  return Number(/name="Box"[\s\S]*?<a:ext cx="(\d+)"/.exec(slide)![1]);
}

afterEach(cleanup);

describe('Page setup slide size', () => {
  test('a preset applies its exact size and scales the content, in one undo step', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <PageSetupDialog open onClose={() => (closed += 1)} />);
    expect((field('size') as unknown as HTMLSelectElement).value).toBe('widescreen');
    expect([field('width').value, field('height').value]).toEqual(['13.33', '7.5']);
    act(() => void fireEvent.change(field('size'), { target: { value: 'standard' } }));
    expect([field('width').value, field('height').value]).toEqual(['10', '7.5']);
    act(() => void fireEvent.click(ok()));
    expect(closed).toBe(1);
    expect([editor.snapshot().slideWidth, editor.snapshot().slideHeight]).toEqual([
      9_144_000, 6_858_000,
    ]);
    expect(boxWidth(editor)).toBe(2_438_400 * 0.75);
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().slideWidth).toBe(12_192_000);
    expect(boxWidth(editor)).toBe(2_438_400);
  });

  test('typed inches give a custom size; portrait swaps them; keep leaves the boxes', async () => {
    const editor = await open();
    withEditor(editor, <PageSetupDialog open onClose={() => {}} />);
    act(() => void fireEvent.change(field('width'), { target: { value: '12' } }));
    expect((field('size') as unknown as HTMLSelectElement).value).toBe('custom');
    act(() => void fireEvent.click(screen.getByRole('radio', { name: 'Portrait' })));
    expect([field('width').value, field('height').value]).toEqual(['7.5', '12']);
    act(() => void fireEvent.click(screen.getByRole('checkbox')));
    act(() => void fireEvent.click(ok()));
    expect([editor.snapshot().slideWidth, editor.snapshot().slideHeight]).toEqual([
      6_858_000, 10_972_800,
    ]);
    expect(boxWidth(editor)).toBe(2_438_400);
  });

  test('a size out of 1 to 56 inches keeps the dialog open with its reason', async () => {
    const editor = await open();
    let closed = 0;
    withEditor(editor, <PageSetupDialog open onClose={() => (closed += 1)} />);
    act(() => void fireEvent.change(field('height'), { target: { value: '0.5' } }));
    act(() => void fireEvent.click(ok()));
    expect(closed).toBe(0);
    expect(screen.getAllByRole('alert').map((a) => a.textContent)).toContain(
      'The slide width and height must each be from 1 to 56 inches.'
    );
    expect(editor.snapshot().slideHeight).toBe(6_858_000);
  });

  test('Custom with the same numbers, or orientation there and back, sends no command', async () => {
    const editor = await open();
    withEditor(editor, <PageSetupDialog open onClose={() => {}} />);
    act(() => void fireEvent.change(field('size'), { target: { value: 'custom' } }));
    act(() => void fireEvent.click(screen.getByRole('radio', { name: 'Portrait' })));
    act(() => void fireEvent.click(screen.getByRole('radio', { name: 'Landscape' })));
    act(() => void fireEvent.click(ok()));
    expect(editor.snapshot().canUndo).toBe(false);
    expect(editor.snapshot().slideWidth).toBe(12_192_000);
  });

  test('viewing mode shows the size as text', async () => {
    const editor = await open();
    act(() => editor.setMode('view'));
    withEditor(editor, <PageSetupDialog open onClose={() => {}} />);
    const section = document.querySelector('[data-pptx-field="slideSize"]')!;
    expect(section.textContent).toContain('Widescreen (16:9)');
    expect(section.textContent).toContain('13.33 × 7.5 in');
  });

  test('OK with untouched fields sends no command', async () => {
    const editor = await open();
    withEditor(editor, <PageSetupDialog open onClose={() => {}} />);
    act(() => void fireEvent.click(ok()));
    expect(editor.snapshot().canUndo).toBe(false);
  });
});
