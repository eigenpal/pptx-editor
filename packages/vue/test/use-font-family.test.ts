// `useFontFamily` in the Vue adapter: the twin of the React test
// (packages/react/src/__tests__/use-font-family.test.tsx), after DOCX Editor's Vue composable of
// the same name. A host builds its own font list from the composable only.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { useFontFamily, type UseFontFamilyReturn } from '../src/index.ts';
import { mountVue } from './toolbar-dom.ts';

const IN = 914_400;

const apps: [App, HTMLElement][] = [];
afterEach(() => {
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
});

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

/** A host font list built from the composable only, as the docs show it. */
function hostFontList(seen: { current: UseFontFamilyReturn | null }) {
  return defineComponent({
    setup() {
      const font = useFontFamily();
      seen.current = font;
      return () =>
        h(
          'ul',
          font.groups.value.map((group) =>
            h(
              'li',
              { key: group.id, 'data-group': group.id },
              group.rows.map((row) =>
                h(
                  'button',
                  {
                    key: row.value,
                    type: 'button',
                    'data-value': row.value,
                    disabled: !font.isEnabled.value || row.disabled,
                    'aria-pressed': font.isCurrent(row),
                    onClick: () => font.setValue(row.value),
                  },
                  row.family
                )
              )
            )
          )
        );
    },
  });
}

describe('useFontFamily (Vue)', () => {
  test('a host list shows the groups, marks the current font, and applies a row', async () => {
    const editor = await open();
    const seen: { current: UseFontFamilyReturn | null } = { current: null };
    const List = hostFontList(seen);
    const { host, app } = await mountVue(editor, () => h(List));
    apps.push([app, host]);
    await nextTick();
    const font = seen.current!;
    expect(font.value.value).toBe('Courier New');
    expect(font.theme.value).toBe(null);
    expect(font.isEnabled.value).toBe(true);
    expect(font.disabledReason.value).toBe(null);
    expect(font.options.value.length).toBeGreaterThan(0);
    expect(font.options.value).toEqual(editor.fontFamilies());
    expect(font.groups.value.map((g) => g.id)).toEqual(['theme', 'document', 'all']);
    expect(font.groups.value[0]!.rows.map((r) => [r.family, r.value])).toEqual([
      ['Georgia', '+mj-lt'],
      ['Verdana', '+mn-lt'],
    ]);
    const groups = [...host.querySelectorAll('[data-group]')].map((g) =>
      g.getAttribute('data-group')
    );
    expect(groups).toEqual(['theme', 'document', 'all']);
    const pressed = [...host.querySelectorAll('[aria-pressed="true"]')].map((b) => b.textContent);
    expect(pressed.includes('Courier New')).toBe(true);
    expect(pressed.includes('Georgia')).toBe(false);

    // A theme row applies the theme font reference, and the mark moves to that row only.
    host.querySelector<HTMLElement>('[data-group="theme"] [data-value="+mj-lt"]')!.click();
    await nextTick();
    await nextTick();
    expect(latinOf(editor)).toBe('+mj-lt');
    expect(font.theme.value).toBe('major');
    const headings = host.querySelector<HTMLElement>('[data-group="theme"] [data-value="+mj-lt"]')!;
    expect(headings.getAttribute('aria-pressed')).toBe('true');
    expect(host.querySelectorAll('[aria-pressed="true"]').length).toBe(1);
  });

  test('outside a root it is disabled with the reason', async () => {
    const seen: { current: UseFontFamilyReturn | null } = { current: null };
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp(hostFontList(seen));
    app.mount(host);
    apps.push([app, host]);
    await nextTick();
    const font = seen.current!;
    expect(font.isEnabled.value).toBe(false);
    expect(typeof font.disabledReason.value).toBe('string');
    expect(font.groups.value.length).toBe(0);
    expect(font.options.value.length).toBe(0);
    // A refusal does nothing and does not throw.
    font.setValue('Arial');
  });
});
