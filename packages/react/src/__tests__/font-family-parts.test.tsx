// The `Toolbar.FontFamily` compound in the React adapter, after DOCX Editor's `FontFamily`
// Trigger / Content / Item parts: a host composes the parts around the shared picker state, an
// item applies its value (a theme font reference keeps the reference), and a part outside the
// compound renders nothing with a development warning.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { Toolbar } from '../index.ts';

const IN = 914_400;

async function open(minorFont = 'Verdana'): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    majorFont: 'Georgia',
    minorFont,
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          text: para(run('Hi', '', '<a:latin typeface="Courier New"/>')),
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  return editor;
}

const latinOf = (editor: PptxEditorInstance): string | undefined => {
  const part = xmlPartOf(editor.session!.store.package, editor.session!.info.slides[0]!.partName)!;
  return /<a:latin typeface="([^"]*)"/.exec(serializeXmlPart(part))?.[1];
};

afterEach(() => {
  cleanup();
  for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
});

describe('Toolbar.FontFamily parts', () => {
  test('a host composes Trigger, Content, and Item around the shared state', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar>
          <Toolbar.FontFamily className="acme-family">
            <Toolbar.FontFamily.Trigger className="acme-trigger">Font</Toolbar.FontFamily.Trigger>
            <Toolbar.FontFamily.Content className="acme-list">
              <Toolbar.FontFamily.Item value="+mj-lt" />
              <Toolbar.FontFamily.Item value="Arial" className="acme-item">
                Plain Arial
              </Toolbar.FontFamily.Item>
              <Toolbar.FontFamily.Item value="Courier New" />
              <Toolbar.FontFamily.Item value={'bad"name'} />
            </Toolbar.FontFamily.Content>
          </Toolbar.FontFamily>
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const trigger = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.textContent).toBe('Font');
    expect(trigger.className).toContain('acme-trigger');
    expect(trigger.className).toContain('acme-family');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    act(() => void fireEvent.click(trigger));
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    // Custom content has no search field, only the host's items.
    expect(document.querySelectorAll('.pptx-font-picker__search').length).toBe(0);
    expect(document.querySelectorAll('.acme-list').length).toBe(1);
    const rows = [...document.querySelectorAll<HTMLElement>('[role="option"]')];
    // The invalid name renders no row; the theme item shows the theme's family and role.
    expect(rows.map((r) => r.textContent)).toEqual([
      'GeorgiaHeadings',
      'Plain Arial',
      'Courier New',
    ]);
    expect(rows[0]!.dataset.theme).toBe('major');
    expect(rows[1]!.className).toContain('acme-item');
    // The current family carries the check mark.
    expect(rows.map((r) => r.getAttribute('aria-selected'))).toEqual(['false', 'false', 'true']);
    act(() => void fireEvent.click(rows[0]!));
    expect(latinOf(editor)).toBe('+mj-lt');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    editor.destroy();
  });

  test('with no children the default trigger and grouped list render', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar>
          <Toolbar.FontFamily className="acme-family" />
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const trigger = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.className).toContain('acme-family');
    expect(trigger.textContent).toContain('Courier New');
    act(() => void fireEvent.click(trigger));
    expect(document.querySelectorAll('.pptx-font-picker__search').length).toBe(1);
    expect(document.querySelectorAll('.pptx-font-picker__group').length).toBe(3);
    editor.destroy();
  });

  test('a host list takes focus at its current option, and the arrow keys move', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    // The theme's body font is empty: its item shows the role, disabled.
    const editor = await open('');
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar>
          <Toolbar.FontFamily>
            <Toolbar.FontFamily.Trigger />
            <Toolbar.FontFamily.Content>
              <Toolbar.FontFamily.Item value="+mn-lt" />
              <Toolbar.FontFamily.Item value="Arial" />
              <Toolbar.FontFamily.Item value="Courier New" />
              <Toolbar.FontFamily.Item value={'bad"name'} />
            </Toolbar.FontFamily.Content>
          </Toolbar.FontFamily>
        </Toolbar>
      </PptxEditorContext.Provider>
    );
    const trigger = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    trigger.focus();
    act(() => void fireEvent.click(trigger));
    const rows = [...document.querySelectorAll<HTMLButtonElement>('[role="option"]')];
    expect(rows.map((r) => r.textContent)).toEqual(['Body', 'Arial', 'Courier New']);
    expect(rows.map((r) => r.disabled)).toEqual([true, false, false]);
    expect(rows[0]!.getAttribute('aria-selected')).toBe('false');
    // The invalid name renders nothing and says why.
    expect(warn.mock.calls.some((c) => String(c[0]).includes('value "bad"name"'))).toBe(true);
    // The current option takes focus; the arrow keys skip the disabled row and wrap.
    expect(document.activeElement).toBe(rows[2]!);
    act(() => void fireEvent.keyDown(rows[2]!, { key: 'ArrowDown' }));
    expect(document.activeElement).toBe(rows[1]!);
    act(() => void fireEvent.keyDown(rows[1]!, { key: 'ArrowUp' }));
    expect(document.activeElement).toBe(rows[2]!);
    // A click on the disabled row applies nothing.
    act(() => void fireEvent.click(rows[0]!));
    expect(latinOf(editor)).toBe('Courier New');
    warn.mockRestore();
    editor.destroy();
  });

  test('a part outside the compound renders nothing and warns', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    const { container } = render(<Toolbar.FontFamily.Item value="Arial" />);
    expect(container.childElementCount).toBe(0);
    expect(
      warn.mock.calls.some((c) =>
        String(c[0]).includes('Toolbar.FontFamily.Item is used outside Toolbar.FontFamily')
      )
    ).toBe(true);
    warn.mockRestore();
  });
});
