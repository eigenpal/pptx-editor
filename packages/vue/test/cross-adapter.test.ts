// The React and Vue toolbars render the same groups, slots, labels, and states for one editor,
// after DOCX Editor's cross-adapter differential test
// (docx-editor/packages/vue/test/cross-adapter-differential.test.ts). Both bars draw every
// registered control of their groups; the test pins the icon command set to the registry, so a
// control missing from either bar fails. `toolbar-controls.test.ts` compares the full markup. The
// tooltip bubble and the grouped font list must also give the same canonical markup. DOCX Editor
// commit adf0e178.

import { afterEach, expect, test } from 'bun:test';
import { createElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef } from 'vue';
import {
  TOOLTIP_DELAY_MS,
  createPptxEditor,
  toolbarGroupsFor,
  type ChromeGroup,
  type PptxEditorInstance,
  type ToolbarContext,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, cell, para, run, shape, table } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { Toolbar } from '../../react/src/Toolbar.tsx';
import { ChromeTooltip as ReactChromeTooltip } from '../../react/src/ChromeTooltip.tsx';
import { PptxEditorToolbar } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { ChromeTooltip } from '../src/ChromeTooltip.tsx';
import { canonicalDom } from './canonical-dom.ts';

/** The icon command controls the preset bar draws for a selection, from the shared registry. */
const iconCommandSlots = (groups: readonly ChromeGroup[]): string[] =>
  groups
    .flatMap((g) =>
      g.controls
        .filter(
          (c) =>
            (c.shape ?? 'icon') === 'icon' &&
            ['command', 'dialog', 'clipboard', 'load', 'replace'].includes(c.state.kind)
        )
        .map((c) => `${g.id}.${c.id}`)
    )
    .sort((a, b) => a.localeCompare(b));

const IN = 914_400;

afterEach(() => cleanup());

interface ButtonState {
  readonly slot: string;
  readonly label: string | null;
  readonly pressed: string | null;
  readonly disabled: boolean;
  readonly tooltip: string | null;
  readonly reason: string | null;
}

/** The icon buttons of a bar, by slot, with what assistive technology and CSS read. */
function buttons(root: Element): ButtonState[] {
  return [...root.querySelectorAll<HTMLButtonElement>('button[data-slot]')]
    .map((b) => ({
      slot: b.dataset.slot ?? '',
      label: b.getAttribute('aria-label'),
      pressed: b.getAttribute('aria-pressed'),
      disabled: b.disabled,
      tooltip: b.getAttribute('data-pptx-tooltip'),
      reason: b.getAttribute('data-pptx-tooltip-reason'),
    }))
    .sort((a, b) => a.slot.localeCompare(b.slot));
}

function groups(root: Element): string[] {
  return [...root.querySelectorAll<HTMLElement>('[role="toolbar"] [role="group"][data-group]')].map(
    (g) => `${g.dataset.group}:${g.getAttribute('aria-label')}`
  );
}

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

test.each<[ToolbarContext]>([['none'], ['shape']])(
  'React and Vue toolbars agree on groups, labels, and icon control state (%s)',
  async (context) => {
    const editor = await open();
    if (context === 'shape')
      editor.setSelection({
        kind: 'shapes',
        nodeIds: [editor.session!.layout(0)!.items[0]!.element.nodeId],
      });
    const ICON_COMMAND_SLOTS = iconCommandSlots(toolbarGroupsFor(context));
    // happy-dom has no layout, so every element is 0 px wide and the React bar would move all
    // its groups into More. A wide bar keeps them all on the bar, as in a browser.
    const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get: () => 4000,
    });
    const reactHost = render(
      createElement(PptxEditorContext.Provider, { value: editor }, createElement(Toolbar))
    ).container;

    const vueHost = document.createElement('div');
    document.body.append(vueHost);
    const app = createApp(
      defineComponent({
        setup() {
          providePptxEditor(shallowRef(editor));
          return () => h(PptxEditorToolbar);
        },
      })
    );
    app.mount(vueHost);
    await nextTick();
    try {
      // Both bars draw the same groups and every control in them, with the same label and state.
      expect(groups(vueHost)).toEqual(groups(reactHost));
      const vue = buttons(vueHost);
      const react = buttons(reactHost);
      const iconSlots = new Set(ICON_COMMAND_SLOTS);
      expect(vue.map((b) => b.slot).filter((s) => iconSlots.has(s))).toEqual(ICON_COMMAND_SLOTS);
      expect(vue).toEqual(react);
    } finally {
      app.unmount();
      vueHost.remove();
      if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
      else delete (HTMLElement.prototype as { clientWidth?: number }).clientWidth;
    }
  }
);

test('in a table, both toolbars add the same short table group', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          columns: [IN, IN],
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: `<a:tr h="${IN}">${cell('A')}${cell('B')}</a:tr>`,
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const frame = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({
    kind: 'text',
    nodeId: frame,
    cell: { row: 0, column: 1 },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
  const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => 8000,
  });
  const reactHost = render(
    createElement(PptxEditorContext.Provider, { value: editor }, createElement(Toolbar))
  ).container;
  const vueHost = document.createElement('div');
  document.body.append(vueHost);
  const app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(PptxEditorToolbar);
      },
    })
  );
  app.mount(vueHost);
  await nextTick();
  try {
    const slots = iconCommandSlots(toolbarGroupsFor('table'));
    const shown = groups(vueHost);
    expect(shown).toEqual(groups(reactHost));
    expect(shown.at(-1)!.split(':')[0]).toBe('table');
    const iconSlots = new Set(slots);
    const react = buttons(reactHost);
    const vue = buttons(vueHost);
    expect(vue.map((b) => b.slot).filter((s) => iconSlots.has(s))).toEqual(slots);
    expect(vue).toEqual(react);
    expect(vue.some((b) => b.slot === 'table.mergeCells')).toBe(true);
    expect(vue.some((b) => b.slot === 'table.headerRow')).toBe(false);
  } finally {
    app.unmount();
    vueHost.remove();
    if (width) Object.defineProperty(HTMLElement.prototype, 'clientWidth', width);
    else delete (HTMLElement.prototype as { clientWidth?: number }).clientWidth;
  }
});

/** The canonical markup of the open tooltip bubble, or of the open font list. */
const canonicalOf = (selector: string): string =>
  canonicalDom(document.querySelector(selector), { self: true, trimText: true });

const wait = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));
const over = (el: Element) =>
  void el.dispatchEvent(new MouseEvent('pointerover', { bubbles: true }));

test('React and Vue draw the same tooltip bubble, with the reason on its own line', async () => {
  const editor = await open();
  const reactHost = render(
    createElement(
      PptxEditorContext.Provider,
      { value: editor },
      createElement('div', { className: 'pptx-editor' }, createElement(Toolbar)),
      createElement(ReactChromeTooltip)
    )
  ).container;
  over(reactHost.querySelector('[data-slot="history.undo"]')!);
  await act(() => wait(TOOLTIP_DELAY_MS + 40));
  const reactBubble = canonicalOf('.pptx-tooltip');
  expect(reactBubble).toContain('pptx-tooltip__reason');
  cleanup();
  // The pointer leaves the removed bar.
  over(document.body);
  expect(document.querySelectorAll('.pptx-tooltip').length).toBe(0);

  const vueHost = document.createElement('div');
  document.body.append(vueHost);
  const app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => [h('div', { class: 'pptx-editor' }, h(PptxEditorToolbar)), h(ChromeTooltip)];
      },
    })
  );
  app.mount(vueHost);
  await nextTick();
  try {
    over(vueHost.querySelector('[data-slot="history.undo"]')!);
    await wait(TOOLTIP_DELAY_MS + 40);
    await nextTick();
    expect(canonicalOf('.pptx-tooltip')).toBe(reactBubble);
  } finally {
    app.unmount();
    vueHost.remove();
  }
});

// A literal family, a theme font reference, and a literal family that has the theme font's name:
// the check mark follows what the text uses, the same in both adapters.
test.each([
  ['Courier New', 'data-family="Courier New"'],
  ['+mj-lt', 'data-theme="major"'],
  ['Georgia', 'data-family="Georgia"'],
])('React and Vue draw the same grouped font list for a run in %s', async (latin, marked) => {
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
          text: para(run('Hi', '', `<a:latin typeface="${latin}"/>`)),
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  editor.setSelection({
    kind: 'shapes',
    nodeIds: [editor.session!.layout(0)!.items[0]!.element.nodeId],
  });
  const reactHost = render(
    createElement(PptxEditorContext.Provider, { value: editor }, createElement(Toolbar))
  ).container;
  act(() => reactHost.querySelector<HTMLElement>('[data-slot="font.family"]')!.click());
  const reactList = canonicalOf('.pptx-font-picker');
  expect(reactList).toContain('pptx-font-picker__role');
  expect(reactList).toContain('--pptx-font-face');
  // Every marked row is the expected one: never a theme row for a literal family.
  const selected = [...document.querySelectorAll('.pptx-font-picker [aria-selected="true"]')];
  expect(selected.length).toBeGreaterThan(0);
  const [name, value] = marked.split('=') as [string, string];
  expect(selected.every((o) => `"${o.getAttribute(name)}"` === value)).toBe(true);
  if (latin !== '+mj-lt') expect(selected.every((o) => !o.hasAttribute('data-theme'))).toBe(true);
  cleanup();
  for (const el of document.querySelectorAll('.pptx-menu')) el.remove();

  const vueHost = document.createElement('div');
  document.body.append(vueHost);
  const app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(PptxEditorToolbar);
      },
    })
  );
  app.mount(vueHost);
  await nextTick();
  try {
    vueHost.querySelector<HTMLElement>('[data-slot="font.family"]')!.click();
    await nextTick();
    await nextTick();
    expect(canonicalOf('.pptx-font-picker')).toBe(reactList);
  } finally {
    app.unmount();
    vueHost.remove();
    for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
  }
});
