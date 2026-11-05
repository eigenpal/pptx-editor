// `useFontFamily` in the React adapter, after DOCX Editor's hook of the same name: the value,
// the offerable families, the packaged picker's groups, and the setter, with no UI, so a host
// builds its own font list. The Vue twin is packages/vue/test/use-font-family.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { useFontFamily, type UseFontFamilyReturn } from '../index.ts';

const IN = 914_400;

afterEach(cleanup);

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    majorFont: 'Georgia',
    minorFont: 'Verdana',
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

/** A host font list built from the hook only, as the docs show it. */
function HostFontList({ seen }: { seen: { current: UseFontFamilyReturn | null } }) {
  const font = useFontFamily();
  seen.current = font;
  return (
    <ul>
      {font.groups.map((group) => (
        <li key={group.id} data-group={group.id}>
          {group.rows.map((row) => (
            <button
              key={row.value}
              type="button"
              data-value={row.value}
              disabled={!font.isEnabled || row.disabled}
              aria-pressed={font.isCurrent(row)}
              onClick={() => font.setValue(row.value)}
            >
              {row.family}
            </button>
          ))}
        </li>
      ))}
    </ul>
  );
}

describe('useFontFamily', () => {
  test('a host list shows the groups, marks the current font, and applies a row', async () => {
    const editor = await open();
    const seen: { current: UseFontFamilyReturn | null } = { current: null };
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <HostFontList seen={seen} />
      </PptxEditorContext.Provider>
    );
    const font = seen.current!;
    expect(font.value).toBe('Courier New');
    expect(font.theme).toBe(null);
    expect(font.isEnabled).toBe(true);
    expect(font.disabledReason).toBe(null);
    expect(font.options.length).toBeGreaterThan(0);
    expect(font.options).toEqual(editor.fontFamilies());
    expect(font.groups.map((g) => g.id)).toEqual(['theme', 'document', 'all']);
    expect(font.groups[0]!.rows.map((r) => [r.family, r.value])).toEqual([
      ['Georgia', '+mj-lt'],
      ['Verdana', '+mn-lt'],
    ]);
    const groups = [...container.querySelectorAll('[data-group]')].map((g) =>
      g.getAttribute('data-group')
    );
    expect(groups).toEqual(['theme', 'document', 'all']);
    const pressed = [...container.querySelectorAll('[aria-pressed="true"]')].map(
      (b) => b.textContent
    );
    expect(pressed.includes('Courier New')).toBe(true);
    expect(pressed.includes('Georgia')).toBe(false);

    // A theme row applies the theme font reference, and the mark moves to that row only.
    const headings = container.querySelector<HTMLElement>(
      '[data-group="theme"] [data-value="+mj-lt"]'
    )!;
    act(() => void fireEvent.click(headings));
    expect(latinOf(editor)).toBe('+mj-lt');
    expect(seen.current!.theme).toBe('major');
    expect(headings.getAttribute('aria-pressed')).toBe('true');
    const marked = container.querySelectorAll('[aria-pressed="true"]').length;
    expect(marked).toBe(1);
  });

  test('outside a root and without a selection it is disabled with the reason', () => {
    const seen: { current: UseFontFamilyReturn | null } = { current: null };
    render(<HostFontList seen={seen} />);
    const font = seen.current!;
    expect(font.isEnabled).toBe(false);
    expect(typeof font.disabledReason).toBe('string');
    expect(font.groups.length).toBe(0);
    expect(font.options.length).toBe(0);
    // A refusal does nothing and does not throw.
    font.setValue('Arial');
  });
});
