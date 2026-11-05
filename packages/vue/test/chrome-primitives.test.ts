// The Vue chrome building blocks render the same markup, ARIA, and keys as their React twins
// (packages/react/src/chrome.tsx, ColorSwatches.tsx, dialog-color.tsx, dialog-form.tsx,
// ContextMenu.tsx, menu-rows.tsx, pane-state.ts), after DOCX Editor's cross-adapter tests.
// Each case renders both adapters and compares a canonical form of the DOM: attributes sorted,
// generated ids replaced by their order of use.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createElement, type ReactElement } from 'react';
import { cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  type ChromeMenuEntry,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import * as R from '../../react/src/chrome.tsx';
import { ColorSwatches as RColorSwatches } from '../../react/src/ColorSwatches.tsx';
import { ColorField as RColorField } from '../../react/src/dialog-color.tsx';
import {
  CheckField as RCheckField,
  FormTabs as RFormTabs,
  NumberField as RNumberField,
  SelectField as RSelectField,
  TextField as RTextField,
} from '../../react/src/dialog-form.tsx';
import { MenuEntries as RMenuEntries } from '../../react/src/menu-rows.tsx';
import { MenuRow as RMenuRow, MenuSubmenu as RMenuSubmenu } from '../../react/src/ContextMenu.tsx';
import {
  Icon,
  MenuItem,
  MenuRow,
  MenuSubmenu,
  Popover,
  PptxEditorRoot,
  ToolButton,
  useNotesPane,
  type UseNotesPaneReturn,
} from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { ColorSwatches } from '../src/ColorSwatches.tsx';
import { ColorField } from '../src/dialog-color.tsx';
import {
  CheckField,
  FormDialog,
  FormTabs,
  NumberField,
  SelectField,
  TextField,
  fieldPart,
  formDialogState,
} from '../src/dialog-form.tsx';
import { createDialogParts, useDialogDraft } from '../src/dialog-parts.ts';
import { MenuEntries, tidySeparators } from '../src/menu-rows.tsx';
import { canonicalDom } from './canonical-dom.ts';

const IN = 914_400;

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

async function mountVue(
  renderFn: () => unknown,
  editor?: PptxEditorInstance | null
): Promise<HTMLElement> {
  app?.unmount();
  host?.remove();
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        if (editor !== undefined) providePptxEditor(shallowRef(editor));
        return renderFn;
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

function mountReact(element: ReactElement, editor?: PptxEditorInstance | null): HTMLElement {
  const wrapped =
    editor === undefined
      ? element
      : createElement(PptxEditorContext.Provider, { value: editor }, element);
  return render(wrapped).container;
}

/** The canonical form of a host's children; a select's reflected `value` is left out. */
const canon = (root: Element): string => canonicalDom(root, { dropSelectValue: true });

async function same(react: ReactElement, vue: () => unknown, editor?: PptxEditorInstance) {
  const r = canon(mountReact(react, editor));
  const v = canon(await mountVue(vue, editor));
  expect(v).toEqual(r);
  return v;
}

async function deck(): Promise<PptxEditorInstance> {
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

const noop = (): void => {};

describe('chrome building blocks', () => {
  test('Icon matches React for a name and for paths', async () => {
    await same(createElement(R.Icon, { name: 'check' }), () => h(Icon, { name: 'check' }));
    await same(createElement(R.Icon, { paths: ['M0 0h10v10z'], className: 'x' }), () =>
      h(Icon, { paths: ['M0 0h10v10z'], className: 'x' })
    );
  });

  test('ToolButton matches React: states, popup, and the spoken disabled reason', async () => {
    const out = await same(
      createElement(R.ToolButton, {
        label: 'Bold',
        icon: 'format_bold',
        onClick: noop,
        pressed: true,
        slot: 'text.bold',
        className: 'acme',
      }),
      () =>
        h(ToolButton, {
          label: 'Bold',
          icon: 'format_bold',
          onClick: noop,
          pressed: true,
          slot: 'text.bold',
          className: 'acme',
        })
    );
    expect(out).toContain('aria-pressed="true"');
    const disabled = await same(
      createElement(R.ToolButton, {
        label: 'Undo',
        icon: 'undo',
        onClick: noop,
        disabled: true,
        disabledReason: 'Nothing to undo',
        hasPopup: 'menu',
      }),
      () =>
        h(ToolButton, {
          label: 'Undo',
          icon: 'undo',
          onClick: noop,
          disabled: true,
          disabledReason: 'Nothing to undo',
          hasPopup: 'menu',
        })
    );
    expect(disabled).toContain('aria-describedby="id0"');
    // One description: the reason text is hidden, so the tree does not show it again as text.
    expect(disabled).toMatch(/<span(?=[^>]*\bid="id0")(?=[^>]*\bhidden)[^>]*>Nothing to undo</);
    expect(disabled).toContain('aria-expanded="false"');
  });

  test('a disabled ToolButton stays focusable and does not run', async () => {
    let ran = 0;
    const el = await mountVue(() =>
      h(ToolButton, { label: 'Undo', onClick: () => (ran += 1), disabled: true })
    );
    const button = el.querySelector('button')!;
    expect(button.disabled).toBe(false);
    button.click();
    expect(ran).toBe(0);
  });

  test('MenuItem matches React as a radio, a checkbox, an option, and a disabled row', async () => {
    const cases: readonly {
      checked?: boolean;
      checkbox?: boolean;
      option?: boolean;
      disabled?: boolean;
      reason?: string;
      title?: string;
    }[] = [
      { checked: true },
      { checked: false, checkbox: true },
      { checked: true, option: true },
      { disabled: true, reason: 'Not now' },
      { title: 'Tip' },
    ];
    for (const extra of cases)
      await same(createElement(R.MenuItem, { onSelect: noop, ...extra, children: 'Row' }), () =>
        h(MenuItem, { onSelect: noop, ...extra }, () => 'Row')
      );
  });
});

describe('Popover', () => {
  async function openPopover(focusAnchor: boolean) {
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const anchor = document.createElement('button');
    anchor.setAttribute('aria-label', 'Align');
    scope.append(anchor);
    document.body.append(scope);
    if (focusAnchor) anchor.focus();
    let closed = 0;
    const open = shallowRef(true);
    await mountVue(() =>
      open.value
        ? h(
            Popover,
            {
              anchor,
              onClose: () => {
                closed += 1;
                open.value = false;
              },
              id: 'pop',
            },
            () => [
              h(MenuItem, { onSelect: noop }, () => 'One'),
              h(MenuItem, { onSelect: noop }, () => 'Two'),
            ]
          )
        : null
    );
    return { scope, anchor, open, closed: () => closed };
  }

  test('renders a menu in the anchor editor scope, named by the trigger', async () => {
    const { scope } = await openPopover(false);
    const panel = scope.querySelector<HTMLElement>('.pptx-menu')!;
    expect(panel !== null).toBe(true);
    expect(panel.getAttribute('role')).toBe('menu');
    expect(panel.getAttribute('aria-label')).toBe('Align');
    expect(panel.id).toBe('pop');
    expect(panel.style.visibility).toBe('');
  });

  test('a keyboard open focuses the first row; arrows move; Escape closes to the trigger', async () => {
    const { scope, anchor, closed } = await openPopover(true);
    const rows = scope.querySelectorAll<HTMLElement>('[role="menuitem"]');
    expect(document.activeElement).toBe(rows[0]!);
    rows[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(rows[1]!);
    rows[1]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(closed()).toBe(1);
    await nextTick();
    expect(scope.querySelector('.pptx-menu') === null).toBe(true);
    expect(document.activeElement).toBe(anchor);
  });

  test('a pointer open focuses the panel; a press outside closes it and focus goes back', async () => {
    const field = document.createElement('input');
    document.body.append(field);
    field.focus();
    const { scope, closed } = await openPopover(false);
    const panel = scope.querySelector<HTMLElement>('.pptx-menu')!;
    // The keys reach the menu, not the field; no ring shows until a key is pressed.
    expect(document.activeElement === panel).toBe(true);
    expect(panel.hasAttribute('data-pointer-open')).toBe(true);
    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(panel.hasAttribute('data-pointer-open')).toBe(false);
    expect(document.activeElement === scope.querySelector('[role="menuitem"]')).toBe(true);
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    expect(closed()).toBe(1);
    await nextTick();
    expect(scope.querySelector('.pptx-menu') === null).toBe(true);
    expect(document.activeElement === field).toBe(true);
    field.remove();
  });

  test('Escape after a pointer open returns focus to the trigger', async () => {
    const field = document.createElement('input');
    document.body.append(field);
    field.focus();
    const { scope, anchor, closed } = await openPopover(false);
    const panel = scope.querySelector<HTMLElement>('.pptx-menu')!;
    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(closed()).toBe(1);
    await nextTick();
    expect(document.activeElement === anchor).toBe(true);
    field.remove();
  });
});

describe('color swatches and the color field', () => {
  test('ColorSwatches matches React, in a menu and in a dialog with a choice', async () => {
    const editor = await deck();
    await same(
      createElement(RColorSwatches, { onPick: noop }),
      () => h(ColorSwatches, { onPick: noop }),
      editor
    );
    const out = await same(
      createElement(RColorSwatches, { onPick: noop, menu: false, selected: '#FF0000' }),
      () => h(ColorSwatches, { onPick: noop, menu: false, selected: '#FF0000' }),
      editor
    );
    expect(out).toContain('aria-pressed="true"');
  });

  test('a theme swatch picks its slot and shift', async () => {
    const editor = await deck();
    const picks: unknown[] = [];
    const el = await mountVue(
      () => h(ColorSwatches, { onPick: (c: unknown) => picks.push(c) }),
      editor
    );
    el.querySelector<HTMLElement>('.pptx-swatches--theme .pptx-swatch')!.click();
    expect(picks[0]).toEqual({ theme: expect.any(String) });
  });

  test('ColorField matches React, and its grid opens, picks, and closes on Escape', async () => {
    const editor = await deck();
    const props = { name: 'color', label: 'Font color', value: '#FF0000', seed: 'keep' } as const;
    await same(
      createElement(RColorField, { ...props, onChange: noop }),
      () => h(ColorField, { ...props, onChange: noop }),
      editor
    );
    const picks: string[] = [];
    const el = await mountVue(
      () => h(ColorField, { ...props, onChange: (p: string) => picks.push(p) }),
      editor
    );
    expect(el.querySelector('select')!.value).toBe('#FF0000');
    const button = el.querySelector<HTMLButtonElement>('.pptx-color-field__button')!;
    button.click();
    await nextTick();
    const panel = el.querySelector<HTMLElement>('.pptx-color-field__panel')!;
    expect(panel !== null).toBe(true);
    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await nextTick();
    expect(el.querySelector('.pptx-color-field__panel') === null).toBe(true);
    button.click();
    await nextTick();
    el.querySelector<HTMLElement>('.pptx-color-field__panel .pptx-swatch')!.click();
    await nextTick();
    expect(picks.length).toBe(1);
  });
});

describe('form dialog fields', () => {
  test('NumberField, TextField, SelectField, CheckField, and FormTabs match React', async () => {
    await same(
      createElement(RNumberField, {
        ...fieldPart('count'),
        name: 'count',
        label: 'Number',
        unit: 'in',
        step: '1',
        min: '1',
        value: '2',
        placeholder: 'Mixed',
        onChange: noop,
      }),
      () =>
        h(NumberField, {
          ...fieldPart('count'),
          name: 'count',
          label: 'Number',
          unit: 'in',
          step: '1',
          min: '1',
          value: '2',
          placeholder: 'Mixed',
          onChange: noop,
        })
    );
    await same(
      createElement(RTextField, {
        name: 't',
        label: 'Text',
        value: 'x',
        font: 'Wingdings',
        preview: '•',
        onChange: noop,
      }),
      () =>
        h(TextField, {
          name: 't',
          label: 'Text',
          value: 'x',
          font: 'Wingdings',
          preview: '•',
          onChange: noop,
        })
    );
    const options = [
      ['a', 'A'],
      ['b', 'B'],
    ] as const;
    await same(
      createElement(RSelectField, {
        name: 's',
        label: 'Pick',
        value: '',
        mixed: 'Mixed',
        options,
        onChange: noop,
      }),
      () =>
        h(SelectField, {
          name: 's',
          label: 'Pick',
          value: '',
          mixed: 'Mixed',
          options,
          onChange: noop,
        })
    );
    await same(
      createElement(RCheckField, { name: 'c', label: 'On', value: true, onChange: noop }),
      () => h(CheckField, { name: 'c', label: 'On', value: true, onChange: noop })
    );
    const tabs = [
      ['font', 'Font'],
      ['spacing', 'Spacing'],
    ] as const;
    await same(
      createElement(RFormTabs, { label: 'Tabs', idBase: 'd', tabs, value: 'font', onChange: noop }),
      () => h(FormTabs, { label: 'Tabs', idBase: 'd', tabs, value: 'font', onChange: noop })
    );
  });

  test('a CheckField with a mixed value is indeterminate', async () => {
    const el = await mountVue(() =>
      h(CheckField, { name: 'c', label: 'On', value: 'mixed', onChange: noop })
    );
    expect(el.querySelector('input')!.indeterminate).toBe(true);
  });

  test('FormTabs move with the arrow keys', async () => {
    const tabs = [
      ['font', 'Font'],
      ['spacing', 'Spacing'],
    ] as const;
    const value = shallowRef('font');
    const el = await mountVue(() =>
      h(FormTabs, {
        label: 'Tabs',
        idBase: 'd',
        tabs,
        value: value.value,
        onChange: (t: string) => (value.value = t),
      })
    );
    el.querySelector<HTMLElement>('[data-tab="font"]')!.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true })
    );
    expect(value.value).toBe('spacing');
  });

  test('FormDialog applies the draft, and a field error marks and focuses its field', async () => {
    interface Fields {
      readonly count: string;
    }
    const parts = createDialogParts<ReturnType<typeof formDialogState<Fields>>, 'count'>(
      'TestDialog'
    );
    const committed: string[] = [];
    let closed = 0;
    const Dialog = defineComponent({
      setup(_, { slots }) {
        const draft = useDialogDraft<Fields>({
          initial: { count: '2' },
          validate: (v) => (Number(v.count) > 0 ? {} : { count: 'Enter a number above 0' }),
          commit: (v) => (committed.push(v.count), { ok: true }),
        });
        const state = formDialogState({ draft, isEnabled: true, onClose: () => (closed += 1) });
        return () =>
          h(
            FormDialog,
            {
              name: 'test',
              title: 'Test',
              okLabel: 'OK',
              cancelLabel: 'Cancel',
              closeLabel: 'Close',
              parts,
              state,
            },
            {
              default: () =>
                h(NumberField, {
                  ...fieldPart('count'),
                  name: 'count',
                  label: 'Number',
                  unit: '',
                  step: '1',
                  value: state.values.value.count,
                  onChange: (v: string) => state.setValue('count', v),
                }),
              host: () => slots.default?.(),
            }
          );
      },
    });
    const el = await mountVue(() =>
      h(Dialog, null, () => [h(parts.Apply, { className: 'acme-ok' })])
    );
    const dialog = el.querySelector<HTMLDialogElement>('dialog[data-pptx-dialog="test"]')!;
    expect(dialog !== null).toBe(true);
    for (const name of ['header', 'title', 'body', 'footer', 'error', 'apply', 'cancel'])
      expect(dialog.querySelector(`[data-pptx-part="${name}"]`) !== null).toBe(true);
    const input = dialog.querySelector<HTMLInputElement>('input[name="count"]')!;
    input.value = '0';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await nextTick();
    const apply = dialog.querySelector<HTMLButtonElement>('[data-pptx-part="apply"]')!;
    expect(apply.classList.contains('acme-ok')).toBe(true);
    apply.click();
    await nextTick();
    await nextTick();
    expect(dialog.querySelector('[data-pptx-part="error"]')!.textContent).toBe(
      'Enter a number above 0'
    );
    expect(input.getAttribute('aria-invalid')).toBe('true');
    expect(dialog.hasAttribute('data-invalid')).toBe(true);
    expect(committed).toEqual([]);
    input.value = '3';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await nextTick();
    apply.click();
    expect(committed).toEqual(['3']);
    expect(closed).toBe(1);
  });
});

describe('menu rows', () => {
  test('MenuRow and MenuSubmenu match React', async () => {
    await same(
      createElement(RMenuRow, {
        label: 'Bold',
        icon: 'format_bold',
        shortcut: '⌘B',
        checked: true,
        slot: 'text.bold',
      }),
      () =>
        h(MenuRow, {
          label: 'Bold',
          icon: 'format_bold',
          shortcut: '⌘B',
          checked: true,
          slot: 'text.bold',
        })
    );
    await same(createElement(RMenuRow, { label: 'Link', href: 'javascript:alert(1)' }), () =>
      h(MenuRow, { label: 'Link', href: 'javascript:alert(1)' })
    );
    await same(createElement(RMenuSubmenu, { label: 'More', icon: 'add', children: null }), () =>
      h(MenuSubmenu, { label: 'More', icon: 'add' }, () => null)
    );
  });

  test('MenuSubmenu opens from the keyboard and closes on ArrowLeft', async () => {
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    document.body.append(scope);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(MenuSubmenu, { label: 'More' }, () => [h(MenuRow, { label: 'Inner' })]),
      }) as Component
    );
    app.mount(scope);
    await nextTick();
    const row = scope.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    row.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    await nextTick();
    await nextTick();
    const panel = scope.querySelector<HTMLElement>('.pptx-menu--submenu')!;
    expect(panel !== null).toBe(true);
    expect(row.getAttribute('aria-expanded')).toBe('true');
    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true }));
    await nextTick();
    expect(scope.querySelector('.pptx-menu--submenu') === null).toBe(true);
    expect(document.activeElement).toBe(row);
  });

  test('MenuEntries of registry rows match React, and a row runs its command', async () => {
    const editor = await deck();
    const entries: ChromeMenuEntry[] = tidySeparators([
      { kind: 'separator' },
      { kind: 'item', slot: 'text.bold' },
      { kind: 'separator' },
      { kind: 'separator' },
      { kind: 'item', slot: 'text.italic' },
      { kind: 'separator' },
    ]);
    expect(entries.map((e) => e.kind)).toEqual(['item', 'separator', 'item']);
    await same(createElement(RMenuEntries, { entries }), () => h(MenuEntries, { entries }), editor);
  });
});

describe('useNotesPane', () => {
  test('the root holds the notes state; toggle and setOpen change it', async () => {
    let pane: UseNotesPaneReturn | null = null;
    const Probe = defineComponent({
      setup() {
        pane = useNotesPane();
        return () => h('span', String(pane!.open.value));
      },
    });
    const el = await mountVue(() => h(PptxEditorRoot, null, () => [h(Probe)]));
    expect(el.textContent).toBe('true');
    pane!.toggle();
    await nextTick();
    expect(el.textContent).toBe('false');
    pane!.toggle();
    pane!.toggle();
    await nextTick();
    expect(el.textContent).toBe('false');
    pane!.setOpen(true);
    await nextTick();
    expect(el.textContent).toBe('true');
  });

  test('outside a root it warns and does nothing', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    let pane: UseNotesPaneReturn | null = null;
    await mountVue(() => {
      return h(
        defineComponent({
          setup() {
            pane = useNotesPane();
            return () => null;
          },
        })
      );
    });
    expect(warn).toHaveBeenCalled();
    pane!.toggle();
    expect(pane!.open.value).toBe(true);
    warn.mockRestore();
  });
});
