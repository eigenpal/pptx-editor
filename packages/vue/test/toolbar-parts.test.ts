import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import {
  createApp,
  createCommentVNode,
  defineComponent,
  h,
  nextTick,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import {
  TOOLBAR_NAMED_PARTS,
  createPptxEditor,
  setNotesTarget,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorToolbar } from '../src/index.ts';
import { providePptxEditor, useEditorCommand } from '../src/composables.ts';
import { tooltipText } from './tooltip-text.ts';

// The Vue twin of the toolbar parts (packages/react/src/__tests__/toolbar-parts.test.tsx).

const IN = 914_400;

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
  // The text box is selected: the bar shows its text, shape, and arrange groups, as in React.
  editor.setSelection({
    kind: 'shapes',
    nodeIds: [editor.session!.layout(0)!.items[0]!.element.nodeId],
  });
  return editor;
}

// Every app a test mounts is unmounted after it, also when a test mounts more than one: a
// toolbar left in the document is found by a later file's `document.querySelector`.
const mounted: [App, HTMLElement][] = [];
afterEach(() => {
  for (const [app, host] of mounted.splice(0)) {
    app.unmount();
    host.remove();
  }
});

async function mount(editor: PptxEditorInstance, render: () => unknown): Promise<HTMLElement> {
  const host = document.createElement('div');
  document.body.append(host);
  const Child = defineComponent({ setup: () => render });
  const app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(Child);
      },
    }) as Component
  );
  app.mount(host);
  mounted.push([app, host]);
  await nextTick();
  return host;
}

const T = PptxEditorToolbar as unknown as Record<string, Component & { pptxSlot?: string }>;

describe('PptxEditorToolbar parts', () => {
  test('every named part stands for the slot in core’s table', () => {
    for (const [name, slot] of Object.entries(TOOLBAR_NAMED_PARTS))
      expect(T[name]!.pptxSlot).toBe(slot);
  });

  test('a named part overrides its slot in place, and hidden removes it', async () => {
    const editor = await open();
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(T.Bold!, { className: 'acme-bold' }),
        h(T.Strike!, { hidden: true }),
      ])
    );
    const bold = el.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(bold.classList.contains('acme-bold')).toBe(true);
    expect(bold.closest('[data-group]')!.getAttribute('data-group')).toBe('text');
    expect(el.querySelector('[data-slot="text.strike"]') === null).toBe(true);
  });

  test('a host group follows the group its after names, with its class', async () => {
    const editor = await open();
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(
          PptxEditorToolbar.Group,
          { id: 'acme', label: 'Acme', after: 'history', class: 'acme' },
          () => [h(PptxEditorToolbar.Action, { label: 'Ask' })]
        ),
      ])
    );
    const ids = [...el.querySelectorAll('.pptx-toolbar > [data-group]')].map((g) =>
      g.getAttribute('data-group')
    );
    expect(ids.slice(0, 3)).toEqual(['history', 'acme', 'font']);
    expect(el.querySelector('[data-group="acme"]')!.classList.contains('acme')).toBe(true);
  });

  test('ToolbarAction emits select, carries state attributes, and keeps the caret', async () => {
    const editor = await open();
    let runs = 0;
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, { preset: false }, () => [
        h(PptxEditorToolbar.Action, { label: 'Send', active: true, onSelect: () => (runs += 1) }),
        h(PptxEditorToolbar.Separator),
        h(PptxEditorToolbar.Action, {
          label: 'Lock',
          disabled: true,
          disabledReason: 'Viewing',
          onSelect: () => (runs += 10),
        }),
      ])
    );
    const [send, lock] = [...el.querySelectorAll<HTMLButtonElement>('button')];
    expect(send!.getAttribute('aria-pressed')).toBe('true');
    expect(send!.hasAttribute('data-active')).toBe(true);
    const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    send!.dispatchEvent(down);
    expect(down.defaultPrevented).toBe(true);
    send!.click();
    lock!.click();
    expect(tooltipText(lock!)).toBe('Lock\nViewing');
    expect(runs).toBe(1);
    expect(el.querySelector('[role="separator"]') !== null).toBe(true);
  });
});

describe('toolbar review fixes (Vue)', () => {
  test('the bar measures for a More menu, a loose action goes after the groups, and a group class applies', async () => {
    const editor = await open();
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(PptxEditorToolbar.Group, { id: 'text', class: 'acme-find' }),
        h(PptxEditorToolbar.Action, { label: 'Send' }),
      ])
    );
    const bar = root.querySelector('.pptx-toolbar')!;
    // As React: the preset bar moves the groups that do not fit into a More menu.
    expect(bar.classList.contains('pptx-toolbar--wrap')).toBe(false);
    expect(bar.getAttribute('data-overflow')).toBe('menu');
    expect(bar.querySelector('[data-group="text"]')!.classList.contains('acme-find')).toBe(true);
    const fixed = bar.querySelector(':scope > [data-fixed]')!;
    expect(fixed.textContent).toBe('Send');
  });

  test('PptxEditorToolbar.Autofit places the Text fitting picker, and hidden removes it', async () => {
    const editor = await open();
    // The default bar leaves it to the Format menu and the Format pane, as in React.
    const preset = await mount(editor, () => h(PptxEditorToolbar));
    expect(preset.querySelector('[data-slot="paragraph.autofit"]') === null).toBe(true);
    const full = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [h(PptxEditorToolbar.Autofit)])
    );
    expect(full.querySelector('[data-slot="paragraph.autofit"]') !== null).toBe(true);
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [h(PptxEditorToolbar.Autofit, { hidden: true })])
    );
    expect(root.querySelector('[data-slot="paragraph.autofit"]') === null).toBe(true);
  });

  test('a named part takes hidden as a typed prop', async () => {
    const editor = await open();
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [h(PptxEditorToolbar.Bold, { hidden: true })])
    );
    expect(root.querySelector('[data-slot="text.bold"]') === null).toBe(true);
  });

  test('ToolbarAction asChild: a span runs on Enter, and misuse falls back to the button', async () => {
    const editor = await open();
    let runs = 0;
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, { preset: false }, () => [
        h(
          PptxEditorToolbar.Action,
          { label: 'Help', asChild: true, onSelect: () => (runs += 1) },
          () => [h('span', null, '?')]
        ),
        h(PptxEditorToolbar.Action, { label: 'Broken', asChild: true }, () => 'text'),
      ])
    );
    const help = root.querySelector<HTMLElement>('[aria-label="Help"]')!;
    expect(help.localName).toBe('span');
    help.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(runs).toBe(1);
    const broken = root.querySelector<HTMLElement>('[aria-label="Broken"]')!;
    expect(broken.localName).toBe('button');
    expect(broken.textContent).toBe('Broken');
  });
});

describe('toolbar review fixes, round 2 (Vue)', () => {
  test('a false v-if adds no empty block, and a link without href runs on Enter', async () => {
    const editor = await open();
    let runs = 0;
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [
        createCommentVNode('v-if', true),
        h(PptxEditorToolbar.Group, { id: 'acme', label: 'Acme' }, () => [
          h(
            PptxEditorToolbar.Action,
            { label: 'Help', asChild: true, onSelect: () => (runs += 1) },
            () => [h('a', null, '?')]
          ),
        ]),
      ])
    );
    expect(root.querySelector('[data-fixed]') === null).toBe(true);
    const help = root.querySelector<HTMLElement>('[aria-label="Help"]')!;
    help.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(runs).toBe(1);
  });

  test('asChild keeps the child attributes, and preventDefault in the child stops the action', async () => {
    const editor = await open();
    let runs = 0;
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, { preset: false }, () => [
        h(
          PptxEditorToolbar.Action,
          { label: 'Send', asChild: true, onSelect: () => (runs += 1) },
          () => [
            h('button', {
              type: 'submit',
              'aria-label': 'Send now',
              onClick: (e: MouseEvent) => e.preventDefault(),
            }),
          ]
        ),
      ])
    );
    const button = root.querySelector<HTMLButtonElement>('button[aria-label="Send now"]')!;
    expect(button.type).toBe('submit');
    button.click();
    expect(runs).toBe(0);
  });
});

describe('toolbar review fixes, round 3 (Vue)', () => {
  test('a named part keeps its control with its slot as icon content, maps class, and leaves empty groups off the bar', async () => {
    const editor = await open();
    const root = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(PptxEditorToolbar.Bold, { class: 'acme-bold' }, () => [h('svg', { class: 'my-icon' })]),
        h(PptxEditorToolbar.Find, { hidden: true }),
        h(PptxEditorToolbar.Replace, { hidden: true }),
        h(PptxEditorToolbar.Group, { id: 'acme', label: 'Acme' }, () => [
          createCommentVNode('v-if', true),
        ]),
      ])
    );
    const bold = root.querySelector('[data-slot="text.bold"]')!;
    expect(bold !== null).toBe(true);
    expect(bold.classList.contains('acme-bold')).toBe(true);
    // As React and DOCX Editor: the default slot of a named button part is its icon.
    expect(bold.querySelector('svg.my-icon') !== null).toBe(true);
    expect(root.querySelector('[data-group="find"]') === null).toBe(true);
    expect(root.querySelector('[data-group="acme"]') === null).toBe(true);
  });
});

describe('toolbar review fixes, round 4', () => {
  test('the bar is one Tab stop, and the arrows move focus, as in React', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const bar = el.querySelector<HTMLElement>('.pptx-toolbar')!;
    const stops = [...bar.querySelectorAll<HTMLButtonElement>('button')].filter(
      (b) => !b.disabled && b.tabIndex === 0
    );
    expect(stops.length).toBe(1);
    const undo = bar.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    undo.focus();
    undo.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, cancelable: true })
    );
    expect(document.activeElement).toBe(bar.querySelector('[data-slot="history.redo"]'));
    expect(undo.tabIndex).toBe(-1);
  });

  test('a ToolbarSlot for a control the preset bar does not draw says so', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      await mount(editor, () =>
        h(PptxEditorToolbar, null, () => [h(T.Slot!, { slotId: 'view.smartGuides' })])
      );
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('"view.smartGuides" is not part of the preset toolbar');
    } finally {
      warn.mockRestore();
    }
  });
});

describe('Vue adapter parts off the preset bar', () => {
  test('a named adapter part (Page setup, Cut) renders its control after the groups', async () => {
    const editor = await open();
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [h(T.PageSetup!), h(T.Cut!)])
    );
    const fixed = el.querySelector('.pptx-toolbar > [data-fixed]')!;
    expect(fixed.querySelector('[data-slot="file.pageSetup"]') !== null).toBe(true);
    expect(fixed.querySelector('[data-slot="edit.cut"]') !== null).toBe(true);
    // Outside a root there is no dialog host: Page setup is off with the reason.
    const setup = fixed.querySelector<HTMLElement>('[data-slot="file.pageSetup"]')!;
    expect(setup.getAttribute('aria-disabled')).toBe('true');
  });
});

describe('toolbar review fixes, round 5', () => {
  const stops = (bar: HTMLElement) =>
    [...bar.querySelectorAll<HTMLButtonElement>('button')].filter(
      (b) => !b.disabled && b.tabIndex === 0
    ).length;

  test('a selection that turns controls on keeps one Tab stop', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const bar = el.querySelector<HTMLElement>('.pptx-toolbar')!;
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    await Promise.resolve();
    expect(stops(bar)).toBe(1);
  });

  test('focus stays on Undo when the last undo turns it off', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    editor.exec({ type: 'toggleBold' });
    await nextTick();
    const undo = el.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    undo.focus();
    undo.click();
    await nextTick();
    const after = el.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(after.getAttribute('aria-disabled')).toBe('true');
    expect(document.activeElement).toBe(after);
    expect(after).toBe(undo);
  });
});

describe('Vue useEditorCommand for adapter slots', () => {
  test('a slot that needs a host is off with the reason; the clipboard runs', async () => {
    const editor = await open();
    const states: Record<string, boolean> = {};
    let reason = null as string | null;
    await mount(editor, () => {
      const dialog = useEditorCommand('file.pageSetup');
      const copy = useEditorCommand('edit.selectAll');
      states.dialog = dialog.isEnabled.value;
      states.selectAll = copy.isEnabled.value;
      reason = dialog.disabledReason.value;
      return h('span');
    });
    expect(states).toEqual({ dialog: false, selectAll: true });
    expect(reason).toBe('disabled.noRoot');
  });
});

describe('Vue bar groups', () => {
  test('the Find group draws its controls, with one separator between groups', async () => {
    const editor = await open();
    // Find is on the bar while nothing is selected.
    editor.setSelection({ kind: 'none' });
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const find = el.querySelector('[data-group="find"]')!;
    expect(find.querySelector('[data-slot="find.find"]') !== null).toBe(true);
    // Replace opens from the find bar, Edit > Replace, and Ctrl/Cmd+H: the bar shows Find only.
    expect(find.querySelector('[data-slot="find.replace"]') === null).toBe(true);
    const seps = el.querySelectorAll('.pptx-toolbar > .pptx-toolbar__sep').length;
    const groups = el.querySelectorAll('.pptx-toolbar > [data-group]').length;
    expect(seps).toBe(Math.max(0, groups - 1));
  });

  test('a named Find part takes icon content; an empty slot replaces the control with nothing', async () => {
    const editor = await open();
    editor.setSelection({ kind: 'none' });
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [h(T.Find!, null, () => [h('i', { class: 'acme-find' })])])
    );
    expect(el.querySelector('[data-slot="find.find"] .acme-find') !== null).toBe(true);
    const el2 = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [h(T.Slot!, { slotId: 'find.find' }, () => [])])
    );
    // As React: slot content (even none) replaces the control.
    expect(el2.querySelector('[data-slot="find.find"]') === null).toBe(true);
  });
});

// A plain `class` on `ToolbarSlot` goes on the control, as React's `className` on
// `Toolbar.Slot` does, in the preset bar and in a hand-built bar.
describe('ToolbarSlot class (Vue)', () => {
  const warnings = (warn: { mock: { calls: unknown[][] } }): string =>
    warn.mock.calls.map((c) => String(c[0])).join('\n');

  test('a plain class and class-name go on the control in the preset bar', async () => {
    const editor = await open();
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(T.Slot!, { slotId: 'text.bold', class: 'acme-bold' }),
        h(T.Slot!, { slotId: 'text.italic', 'class-name': 'acme-italic' }),
        h(T.Slot!, { slotId: 'text.underline', class: 'acme-a', className: 'acme-b' }),
        h(T.Strike!, { class: ['acme-strike', { 'acme-on': true }] }),
      ])
    );
    const control = (slot: string) => el.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
    expect(control('text.bold').classList.contains('acme-bold')).toBe(true);
    expect(control('text.bold').classList.contains('pptx-toolbar__button')).toBe(true);
    expect(control('text.bold').closest('[data-group]')!.getAttribute('data-group')).toBe('text');
    expect(control('text.italic').classList.contains('acme-italic')).toBe(true);
    expect([...control('text.underline').classList]).toEqual(
      expect.arrayContaining(['acme-a', 'acme-b'])
    );
    expect([...control('text.strike').classList]).toEqual(
      expect.arrayContaining(['acme-strike', 'acme-on'])
    );
  });

  test('a plain class and className go on the control in a hand-built bar', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      const el = await mount(editor, () =>
        h(PptxEditorToolbar, { preset: false }, () => [
          h(T.Slot!, { slotId: 'text.bold', class: 'acme-bold' }),
          h(T.Slot!, { slotId: 'text.italic', className: 'acme-italic' }),
        ])
      );
      const bold = el.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
      expect(bold.classList.contains('acme-bold')).toBe(true);
      expect(el.querySelector('[data-slot="text.italic"]')!.classList.contains('acme-italic')).toBe(
        true
      );
      // The class lands once, on the control: not dropped with a fragment warning.
      expect(el.querySelectorAll('.acme-bold')).toHaveLength(1);
      expect(warnings(warn)).not.toContain('Extraneous non-props attributes');
    } finally {
      warn.mockRestore();
    }
  });

  test('a class on a slot whose content replaces the control warns, in both bars', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const editor = await open();
      const el = await mount(editor, () => [
        h(PptxEditorToolbar, null, () => [
          h(T.Slot!, { slotId: 'text.bold', class: 'acme-bold' }, () => [h('i', { id: 'mine' })]),
        ]),
        h(PptxEditorToolbar, { preset: false }, () => [
          h(T.Slot!, { slotId: 'text.italic', className: 'acme-italic' }, () => [
            h('i', { id: 'mine-2' }),
          ]),
        ]),
      ]);
      expect(el.querySelector('#mine')).not.toBeNull();
      expect(el.querySelector('#mine-2')).not.toBeNull();
      expect(el.querySelector('.acme-bold')).toBeNull();
      const text = warnings(warn);
      expect(text).toContain(
        'ToolbarSlot "text.bold" has its own content, which replaces the control, so its class is ignored.'
      );
      expect(text).toContain('ToolbarSlot "text.italic" has its own content');
    } finally {
      warn.mockRestore();
    }
  });
});

// Only the class reaches the control. Any other attribute on a slot part (`style`, `id`,
// `data-*`, `aria-*`, a listener) warns in both bars, as `Toolbar.Slot` does in React, and
// never falls through to Vue's own "Extraneous non-props attributes" warning.
describe('ToolbarSlot attributes the control does not take (Vue)', () => {
  const spy = () => spyOn(console, 'warn').mockImplementation(() => {});
  const text = (warn: ReturnType<typeof spy>): string =>
    warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');

  for (const preset of [true, false]) {
    test(`data-*, style, and id warn and stay off the control (preset ${preset})`, async () => {
      resetDevWarnings();
      const warn = spy();
      try {
        const editor = await open();
        const el = await mount(editor, () =>
          h(PptxEditorToolbar, { preset }, () => [
            h(T.Slot!, { slotId: 'text.bold', 'data-test': 'bold', style: 'color: red' }),
            h(T.Italic!, { id: 'acme-italic', class: 'acme-italic' }),
          ])
        );
        const bold = el.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
        expect(bold.hasAttribute('data-test')).toBe(false);
        expect(el.querySelector('#acme-italic')).toBeNull();
        expect(el.querySelector('[data-slot="text.italic"]')!.classList).toContain('acme-italic');
        const said = text(warn);
        expect(said).toContain(
          'ToolbarSlot "text.bold" does not pass data-test, style to its control, so they are ignored.'
        );
        expect(said).toContain(
          'PptxEditorToolbar.Italic "text.italic" does not pass id to its control, so it is ignored.'
        );
        expect(said).not.toContain('Extraneous non-props attributes');
      } finally {
        warn.mockRestore();
      }
    });
  }

  test('a named part takes icon in the preset bar, with no warning', async () => {
    resetDevWarnings();
    const warn = spy();
    try {
      const editor = await open();
      const el = await mount(editor, () =>
        h(PptxEditorToolbar, null, () => [h(T.Bold!, { icon: h('i', { id: 'acme-icon' }) })])
      );
      expect(el.querySelector('[data-slot="text.bold"] #acme-icon')).not.toBeNull();
      expect(text(warn)).not.toContain('text.bold');
    } finally {
      warn.mockRestore();
    }
  });
});

describe('the Vue bar for the selection', () => {
  const groups = (el: HTMLElement) =>
    [...el.querySelectorAll('.pptx-toolbar > [data-group]')].map((g) =>
      g.getAttribute('data-group')
    );

  test('Delete removes its own group; the focus stays in the bar, as in React', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const remove = el.querySelector<HTMLElement>('[data-slot="more.delete"]')!;
    remove.focus();
    remove.click();
    await nextTick();
    await nextTick();
    expect(groups(el)).toEqual(['history', 'insert', 'find']);
    expect(el.querySelector('.pptx-toolbar')!.contains(document.activeElement)).toBe(true);
  });

  test('focus the user moves off the bar stays off when the groups change, as in React', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const remove = el.querySelector<HTMLElement>('[data-slot="more.delete"]')!;
    remove.focus();
    remove.blur();
    await Promise.resolve();
    editor.setSelection({ kind: 'none' });
    await nextTick();
    await nextTick();
    expect(document.activeElement).toBe(document.body);
  });

  test('contextual false draws a placed slide slot after the groups, as in React', async () => {
    const editor = await open();
    const el = await mount(editor, () =>
      h(PptxEditorToolbar, { contextual: false }, () => [
        h(PptxEditorToolbar.Slot, { slotId: 'slide.background' }),
      ])
    );
    expect(el.querySelector('[data-slot="slide.background"]') !== null).toBe(true);
  });

  test('a focused notes field shows Undo, Redo, the notes formats, and Find, as in React', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorToolbar));
    const applied: string[] = [];
    setNotesTarget(editor, {
      formats: { bold: false, italic: false, underline: false, strike: false },
      readOnly: false,
      canFormat: true,
      apply: (format) => (applied.push(format), true),
    });
    await nextTick();
    await nextTick();
    try {
      const slots = [...el.querySelectorAll('.pptx-toolbar [data-slot]')].map((e) =>
        e.getAttribute('data-slot')
      );
      expect(slots).toEqual([
        'history.undo',
        'history.redo',
        'text.bold',
        'text.italic',
        'text.underline',
        'text.strike',
        'find.find',
      ]);
      el.querySelector<HTMLElement>('[data-slot="text.strike"]')!.click();
      expect(applied).toEqual(['strike']);
    } finally {
      setNotesTarget(editor, null);
    }
  });

  test('contextual false keeps every group, as in React', async () => {
    const editor = await open();
    editor.setSelection({ kind: 'none' });
    const el = await mount(editor, () => h(PptxEditorToolbar, { contextual: false }));
    expect(groups(el)).toContain('font');
    expect(groups(el)).toContain('shape');
  });
});
