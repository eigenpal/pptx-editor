// The Animations tab of the side panel in the React adapter: the labels of read-only effects,
// distinct names for the advance controls, an error for a refused time, and a chosen effect
// that stays on its own slide. `packages/vue/test/animations-pane.test.ts` checks that the Vue
// twin renders the same markup.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../../scripts/fixtures/timing.ts';
import { AnimationsPane } from '../AnimationsPane.tsx';
import { PptxEditorContext } from '../context.ts';

const IN = 914_400;
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

afterEach(() => cleanup());

const box = (id: number, name: string, x: number) =>
  shape({ id, name, x: x * IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(name)) });

async function open(...exts: string[]): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: exts.map((ext) => ({
      layout: 'blank' as const,
      shapes: box(2, 'A', 1) + box(3, 'B', 4),
      ...(ext ? { ext } : {}),
    })),
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

function mount(editor: PptxEditorInstance): HTMLElement {
  return render(
    <PptxEditorContext.Provider value={editor}>
      <AnimationsPane />
    </PptxEditorContext.Provider>
  ).container;
}

const flush = () => act(async () => undefined);

describe('Animations pane (React)', () => {
  test('a read-only main sequence shows the trigger of each effect', async () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:p14="${P14}" Requires="p14">` +
      clickEntrances([2]) +
      '</mc:Choice><mc:Fallback>' +
      clickEntrances([2]) +
      '</mc:Fallback></mc:AlternateContent>';
    const el = mount(await open(wrapped));
    const row = el.querySelector('.pptx-animations__row')!;
    expect(row.textContent).toContain('On click');
    expect(row.textContent).not.toContain('Missing shape');
  });

  test('the advance controls have different names', async () => {
    const el = mount(await open(''));
    const names = [...el.querySelectorAll('input')].map(
      (i) => i.getAttribute('aria-label') ?? i.closest('label')?.textContent
    );
    expect(names).toContain('Advance after');
    expect(names).toContain('Advance automatically');
    expect(new Set(names).size).toBe(names.length);
  });

  test('a refused time stays in its field with an error', async () => {
    const editor = await open(clickEntrances([2]));
    const el = mount(editor);
    fireEvent.click(el.querySelector('.pptx-animations__row')!);
    await flush();
    const delay = el.querySelector<HTMLInputElement>('input[aria-label="Delay"]')!;
    // 4,000 s is past the longest delay.
    fireEvent.change(delay, { target: { value: '4000' } });
    fireEvent.blur(delay);
    await flush();
    expect(delay.getAttribute('aria-invalid')).toBe('true');
    expect(delay.value).toBe('4000');
    const error = el.querySelector(`#${CSS.escape(delay.getAttribute('aria-describedby')!)}`);
    expect(error?.getAttribute('role')).toBe('alert');
    expect(editor.snapshot().animations!.effects[0].delayMs).toBe(0);
    // A good value clears the error.
    fireEvent.change(delay, { target: { value: '1' } });
    fireEvent.blur(delay);
    await flush();
    expect(delay.getAttribute('aria-invalid')).toBe('false');
    expect(el.querySelector('[role="alert"]')).toBeNull();
    expect(editor.snapshot().animations!.effects[0].delayMs).toBe(1000);
  });

  test('the chosen effect stays on its slide', async () => {
    // Both slides have an effect with the same id.
    const editor = await open(clickEntrances([2]), clickEntrances([3]));
    const el = mount(editor);
    fireEvent.click(el.querySelector('.pptx-animations__row')!);
    await flush();
    expect(el.querySelector('.pptx-animations__fields')).not.toBeNull();
    await act(async () => void editor.goToSlide(1));
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(el.querySelector('.pptx-animations__fields')).toBeNull();
    expect(el.querySelector('[aria-pressed="true"]')).toBeNull();
  });
});
