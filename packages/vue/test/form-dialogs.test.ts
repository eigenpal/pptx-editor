// The Vue twins of the form dialogs (Paragraph, Font, Bullets and Numbering, Columns) and the
// Background dialog: the same markup as the React dialogs on one deck, the same command on OK as
// one undo step, the same field errors, the named parts, and the popups registry.

import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { createElement, type ReactElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import {
  createApp,
  defineComponent,
  h,
  nextTick,
  reactive,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import {
  createPptxEditor,
  requestDialogSection,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { BackgroundDialog as RBackgroundDialog } from '../../react/src/BackgroundDialog.tsx';
import { BulletsAndNumberingDialog as RBulletsDialog } from '../../react/src/BulletsAndNumberingDialog.tsx';
import { ColumnsDialog as RColumnsDialog } from '../../react/src/ColumnsDialog.tsx';
import { FontDialog as RFontDialog } from '../../react/src/FontDialog.tsx';
import { ParagraphDialog as RParagraphDialog } from '../../react/src/ParagraphDialog.tsx';
import { providePptxEditor } from '../src/composables.ts';
import {
  BackgroundDialog,
  BulletsAndNumberingDialog,
  ColumnsDialog,
  FontDialog,
  ParagraphDialog,
  PptxEditorRoot,
  definePopup,
  useColumnsDialog,
  useEditorDialog,
  type PptxEditorDialogKey,
  type PptxEditorPopups,
} from '../src/index.ts';
import { canonicalDom } from './canonical-dom.ts';

const IN = 914_400;

// An inactive tab panel is `visibility: hidden` (editor.css), so a browser does not focus a
// control in it. happy-dom computes no CSS: refuse that focus here, as a browser does.
const nativeFocus = HTMLElement.prototype.focus;
beforeAll(() => {
  HTMLElement.prototype.focus = function focus(this: HTMLElement, options?: FocusOptions) {
    if (this.closest('.pptx-dialog__panels > [data-inactive]')) return;
    nativeFocus.call(this, options);
  };
});
afterAll(() => {
  HTMLElement.prototype.focus = nativeFocus;
});

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

async function deck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text:
            para(run('one'), '<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>') +
            para(run('two')),
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

/** Select both paragraphs of the text box. */
function selectText(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({
    kind: 'text',
    nodeId,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 1, offset: 3 },
  });
}

/** Select the text box as a shape. */
function selectShape(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
}

async function mountVue(renderFn: () => unknown, editor: PptxEditorInstance): Promise<HTMLElement> {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return renderFn;
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  await nextTick();
  return host;
}

function mountReact(element: ReactElement, editor: PptxEditorInstance): HTMLElement {
  let container: HTMLElement | null = null;
  act(() => {
    container = render(createElement(PptxEditorContext.Provider, { value: editor }, element))
      .container as HTMLElement;
  });
  return container!;
}

/**
 * The canonical form of a dialog host. Form state is a property, not an attribute, so it is
 * compared as the user sees it. `autofocus` is a DOM attribute in Vue and a mount-time focus call
 * in React; the initial-focus test checks the focus itself.
 */
const canon = (root: Element): string =>
  canonicalDom(root, {
    dropSelectValue: true,
    splitIdReferences: true,
    dropAttributes: ['autofocus'],
    formState: true,
  });

/** The focused control, named by its tag and `name`, or `body`. */
const focused = (): string => {
  const el = document.activeElement;
  return el && el !== document.body
    ? `${el.tagName.toLowerCase()}[name=${el.getAttribute('name') ?? ''}]`
    : 'body';
};

const dialogOf = (name: string) =>
  document.querySelector<HTMLDialogElement>(`dialog[data-pptx-dialog="${name}"]`)!;
const field = (dialog: Element, name: string) =>
  dialog.querySelector<HTMLInputElement & HTMLSelectElement>(`[name="${name}"]`)!;
const part = (dialog: Element, name: string) =>
  dialog.querySelector<HTMLElement>(`[data-pptx-part="${name}"]`)!;

function type(el: HTMLInputElement, value: string): void {
  el.value = value;
  el.dispatchEvent(new Event('input', { bubbles: true }));
}
function choose(el: HTMLSelectElement, value: string): void {
  el.value = value;
  el.dispatchEvent(new Event('change', { bubbles: true }));
}

interface Pair {
  readonly name: string;
  readonly react: (props: { open: boolean; onClose: () => void }) => ReactElement;
  readonly vue: Component;
  readonly select: (editor: PptxEditorInstance) => void;
}

const PAIRS: readonly Pair[] = [
  {
    name: 'paragraph',
    react: (p) => createElement(RParagraphDialog, p),
    vue: ParagraphDialog as Component,
    select: selectText,
  },
  {
    name: 'font',
    react: (p) => createElement(RFontDialog, p),
    vue: FontDialog as Component,
    select: selectText,
  },
  {
    name: 'bulletsAndNumbering',
    react: (p) => createElement(RBulletsDialog, p),
    vue: BulletsAndNumberingDialog as Component,
    select: selectText,
  },
  {
    name: 'columns',
    react: (p) => createElement(RColumnsDialog, p),
    vue: ColumnsDialog as Component,
    select: selectShape,
  },
  {
    name: 'background',
    react: (p) => createElement(RBackgroundDialog, p),
    vue: BackgroundDialog as Component,
    select: () => undefined,
  },
];

describe('the React and Vue form dialogs', () => {
  for (const pair of PAIRS)
    test(`${pair.name}: the same markup on one deck`, async () => {
      const editor = await deck();
      pair.select(editor);
      const react = canon(mountReact(pair.react({ open: true, onClose: () => {} }), editor));
      const reactFocus = focused();
      expect(dialogOf(pair.name).contains(document.activeElement)).toBe(true);
      cleanup();
      const vue = canon(
        await mountVue(() => h(pair.vue, { open: true, onClose: () => {} }), editor)
      );
      expect(vue).toContain(`data-pptx-dialog="${pair.name}"`);
      expect(vue).toEqual(react);
      // `canon` leaves out `autofocus`: the initial focus is compared here instead.
      expect(focused()).toBe(reactFocus);
      expect(dialogOf(pair.name).contains(document.activeElement)).toBe(true);
    });

  test('the first field takes focus when showModal focused the <dialog> first, in both adapters', async () => {
    // A browser's showModal focuses the <dialog> itself while its fields are still empty.
    const proto = HTMLDialogElement.prototype as { showModal?: () => void };
    const real = Object.getOwnPropertyDescriptor(proto, 'showModal');
    proto.showModal = function (this: HTMLDialogElement) {
      this.setAttribute('open', '');
      this.tabIndex = -1;
      this.focus();
    };
    try {
      const editor = await deck();
      selectText(editor);
      mountReact(createElement(RParagraphDialog, { open: true, onClose: () => {} }), editor);
      const reactFocus = focused();
      expect(reactFocus).toMatch(/^(input|select)\[name=/);
      cleanup();
      await mountVue(() => h(ParagraphDialog, { open: true, onClose: () => {} }), editor);
      expect(focused()).toBe(reactFocus);
    } finally {
      if (real) Object.defineProperty(proto, 'showModal', real);
      else delete proto.showModal;
    }
  });

  test('the Font dialog opens on the tab a menu row asks for, in both adapters', async () => {
    const editor = await deck();
    selectText(editor);
    requestDialogSection(editor, 'font', 'spacing');
    const react = canon(
      mountReact(createElement(RFontDialog, { open: true, onClose: () => {} }), editor)
    );
    cleanup();
    requestDialogSection(editor, 'font', 'spacing');
    const vue = canon(
      await mountVue(() => h(FontDialog, { open: true, onClose: () => {} }), editor)
    );
    expect(vue).toEqual(react);
    expect(dialogOf('font').querySelector('[role="tab"][aria-selected="true"]')!.textContent).toBe(
      'Character spacing'
    );
  });

  test('Bullets: a list whose levels differ opens on Bulleted, in both adapters', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const level = (lvl: number, char: string) =>
      `<a:lvl${lvl}pPr marL="${lvl * 457200}" indent="-228600"><a:buFont typeface="Arial"/>` +
      `<a:buChar char="${char}"/></a:lvl${lvl}pPr>`;
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: 2 * IN,
            lstStyle: level(1, '•') + level(2, '–'),
            text: para(run('one')) + para(run('two'), '<a:pPr lvl="1"/>'),
          }),
        },
      ],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    selectText(editor);
    // The Numbering gallery's row asks for Numbered; the paragraphs are all bulleted.
    requestDialogSection(editor, 'bulletsAndNumbering', 'number');
    const react = canon(
      mountReact(createElement(RBulletsDialog, { open: true, onClose: () => {} }), editor)
    );
    cleanup();
    requestDialogSection(editor, 'bulletsAndNumbering', 'number');
    const vue = canon(
      await mountVue(() => h(BulletsAndNumberingDialog, { open: true, onClose: () => {} }), editor)
    );
    expect(vue).toEqual(react);
    const dialog = dialogOf('bulletsAndNumbering');
    expect(dialog.querySelector('[role="tab"][aria-selected="true"]')!.textContent).toBe(
      'Bulleted'
    );
    expect([field(dialog, 'sizePercent').value, field(dialog, 'startAt').value]).toEqual([
      '100',
      '1',
    ]);
  });

  test('closed, a dialog renders nothing; the same mount opened renders it', async () => {
    const editor = await deck();
    selectShape(editor);
    const state = reactive({ open: false });
    await mountVue(() => h(ColumnsDialog, { open: state.open, onClose: () => {} }), editor);
    expect(document.querySelectorAll('dialog').length).toBe(0);
    state.open = true;
    await nextTick();
    expect(document.querySelectorAll('dialog[data-pptx-dialog="columns"]').length).toBe(1);
  });
});

describe('form dialog behavior', () => {
  async function open(
    dialog: Component,
    editor: PptxEditorInstance,
    children?: () => unknown
  ): Promise<{ closed: number }> {
    const state = { closed: 0 };
    await mountVue(
      () => h(dialog, { open: true, onClose: () => (state.closed += 1) }, children),
      editor
    );
    return state;
  }

  test('Columns: OK sends one setColumns command, one undo step', async () => {
    const editor = await deck();
    selectShape(editor);
    const state = await open(ColumnsDialog as Component, editor);
    const dialog = dialogOf('columns');
    type(field(dialog, 'count'), '3');
    await nextTick();
    part(dialog, 'apply').click();
    await nextTick();
    expect(state.closed).toBe(1);
    expect(editor.snapshot().format?.columns?.count).toBe(3);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().format?.columns?.count ?? 1).toBe(1);
  });

  test('Columns: a value out of range names the field, marks it, and any edit clears it', async () => {
    const editor = await deck();
    selectShape(editor);
    const state = await open(ColumnsDialog as Component, editor);
    const dialog = dialogOf('columns');
    type(field(dialog, 'count'), '40');
    await nextTick();
    part(dialog, 'apply').click();
    await nextTick();
    await nextTick();
    expect(state.closed).toBe(0);
    expect(part(dialog, 'error').textContent).toContain('Number must be');
    expect(field(dialog, 'count').getAttribute('aria-invalid')).toBe('true');
    expect(dialog.hasAttribute('data-invalid')).toBe(true);
    type(field(dialog, 'spacing'), '0.2');
    await nextTick();
    await nextTick();
    expect(part(dialog, 'error').textContent).toBe('');
    expect(field(dialog, 'count').hasAttribute('aria-invalid')).toBe(false);
  });

  test('Paragraph: OK with no change only closes; a change applies as one step', async () => {
    const editor = await deck();
    selectText(editor);
    const state = await open(ParagraphDialog as Component, editor);
    const dialog = dialogOf('paragraph');
    choose(field(dialog, 'alignment'), 'center');
    await nextTick();
    choose(field(dialog, 'special'), 'firstLine');
    await nextTick();
    // Picking a special indent fills its amount, as in React.
    expect(field(dialog, 'specialBy').disabled).toBe(false);
    expect(field(dialog, 'specialBy').value).toBe('0.5');
    part(dialog, 'apply').click();
    await nextTick();
    expect(state.closed).toBe(1);
    expect(editor.snapshot().format?.paragraph?.alignment).toBe('center');
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().format?.paragraph?.alignment).not.toBe('center');
  });

  test('Font: an error on the hidden tab shows that tab; a tab change clears it', async () => {
    const editor = await deck();
    selectText(editor);
    await open(FontDialog as Component, editor);
    const dialog = dialogOf('font');
    const selectedTab = () =>
      dialog.querySelector('[role="tab"][aria-selected="true"]')!.getAttribute('data-tab');
    (dialog.querySelector('[data-tab="spacing"]') as HTMLElement).click();
    await nextTick();
    choose(field(dialog, 'spacing'), 'expanded');
    await nextTick();
    (dialog.querySelector('[data-tab="font"]') as HTMLElement).click();
    await nextTick();
    // 2000 pt is beyond the engine's limit for character spacing.
    type(field(dialog, 'by'), '2000');
    await nextTick();
    part(dialog, 'apply').click();
    await nextTick();
    await nextTick();
    expect(selectedTab()).toBe('spacing');
    expect(part(dialog, 'error').textContent).not.toBe('');
    expect(field(dialog, 'by').getAttribute('aria-invalid')).toBe('true');
    (dialog.querySelector('[data-tab="font"]') as HTMLElement).click();
    await nextTick();
    await nextTick();
    expect(part(dialog, 'error').textContent).toBe('');
  });

  test('Font: Current color shows the text color, and Offset follows both script boxes', async () => {
    const editor = await deck();
    selectText(editor);
    await open(FontDialog as Component, editor);
    const dialog = dialogOf('font');
    const color = editor.snapshot().format?.color;
    expect(typeof color).toBe('string');
    expect(field(dialog, 'color').value).toBe('keep');
    const swatch = field(dialog, 'color')
      .closest('.pptx-color-field')!
      .querySelector<HTMLElement>('.pptx-color-field__button')!;
    expect(swatch.getAttribute('data-color')).toBe(color!);
    const order = [...dialog.querySelectorAll('[name]')].map((el) => el.getAttribute('name'));
    expect(order.indexOf('superscript')).toBeLessThan(order.indexOf('subscript'));
    expect(order.indexOf('subscript')).toBeLessThan(order.indexOf('offset'));
  });

  test('Font: a checkbox pair stays exclusive, and OK applies it', async () => {
    const editor = await deck();
    selectText(editor);
    await open(FontDialog as Component, editor);
    const dialog = dialogOf('font');
    field(dialog, 'superscript').click();
    await nextTick();
    expect(field(dialog, 'offset').disabled).toBe(false);
    field(dialog, 'subscript').click();
    await nextTick();
    expect(field(dialog, 'superscript').checked).toBe(false);
    expect(field(dialog, 'subscript').checked).toBe(true);
    part(dialog, 'apply').click();
    await nextTick();
    expect(editor.snapshot().format?.font?.baseline ?? 0).toBeLessThan(0);
  });

  test('Bullets: the Numbered tab names its own size field in an error', async () => {
    const editor = await deck();
    selectText(editor);
    await open(BulletsAndNumberingDialog as Component, editor);
    const dialog = dialogOf('bulletsAndNumbering');
    (dialog.querySelector('[data-tab="number"]') as HTMLElement).click();
    await nextTick();
    choose(field(dialog, 'numberStyle'), 'arabicPeriod');
    await nextTick();
    type(field(dialog, 'numberSizePercent'), '900');
    await nextTick();
    part(dialog, 'apply').click();
    await nextTick();
    await nextTick();
    expect(field(dialog, 'numberSizePercent').getAttribute('aria-invalid')).toBe('true');
    type(field(dialog, 'numberSizePercent'), '100');
    await nextTick();
    part(dialog, 'apply').click();
    await nextTick();
    expect(editor.snapshot().format?.bulletFormat?.kind).toBe('number');
  });

  test('Background: a swatch pick applies as one step; All slides needs a pick first', async () => {
    const editor = await deck();
    const state = await open(BackgroundDialog as Component, editor);
    const dialog = dialogOf('background');
    const allSlides = dialog.querySelector<HTMLInputElement>(
      '[data-pptx-field="allSlides"] input'
    )!;
    // The slide follows the layout: the draft holds the layout's background, which All slides
    // may apply.
    allSlides.click();
    await nextTick();
    const swatch = dialog.querySelector<HTMLButtonElement>(
      '[data-pptx-field="solid"] .pptx-swatch'
    )!;
    swatch.click();
    await nextTick();
    expect(swatch.getAttribute('aria-pressed')).toBe('true');
    part(dialog, 'apply').click();
    await nextTick();
    expect(state.closed).toBe(1);
    expect(editor.snapshot().slideBackground.inherited).toBe(false);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slideBackground.inherited).toBe(true);
  });

  test('Background: Reset picks the layout background and the status line says so', async () => {
    const editor = await deck();
    await open(BackgroundDialog as Component, editor);
    const dialog = dialogOf('background');
    dialog.querySelector<HTMLButtonElement>('[data-pptx-field="gradient"] .pptx-swatch')!.click();
    await nextTick();
    expect(dialog.querySelector('[role="status"]')!.textContent).toBe('');
    part(dialog, 'reset').click();
    await nextTick();
    expect(dialog.querySelector('[role="status"]')!.textContent).not.toBe('');
  });

  test('Background: Custom gradient opens the gradient editor inside the dialog', async () => {
    const editor = await deck();
    await open(BackgroundDialog as Component, editor);
    const dialog = dialogOf('background');
    const custom = [...dialog.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === 'Custom gradient…'
    )!;
    custom.click();
    await nextTick();
    await nextTick();
    const gradient = dialog.querySelector<HTMLDialogElement>('[data-pptx-dialog="gradient"]')!;
    expect(gradient.querySelector('.pptx-gradient__bar') !== null).toBe(true);
    // A gradient that no preset shows: the status line names it.
    const angle = gradient.querySelector<HTMLInputElement>('input[type="number"]')!;
    type(angle, '45');
    await nextTick();
    [...gradient.querySelectorAll<HTMLButtonElement>('.pptx-dialog__footer button')]
      .at(-1)!
      .click();
    await nextTick();
    expect(dialog.querySelector('[role="status"]')!.textContent).not.toBe('');
  });

  test('named parts: a hidden Field, a host class on Apply, and the draft hook', async () => {
    const editor = await deck();
    selectShape(editor);
    let read = '';
    const Probe = defineComponent({
      setup() {
        const state = useColumnsDialog();
        return () => h('i', { 'data-probe': '' }, (read = state.values.value.count));
      },
    });
    await open(ColumnsDialog as Component, editor, () => [
      h(ColumnsDialog.Field, { name: 'spacing', hidden: true }),
      h(ColumnsDialog.Apply, { className: 'acme-ok' }),
      h(Probe),
    ]);
    const dialog = dialogOf('columns');
    expect(dialog.querySelector('[name="spacing"]') === null).toBe(true);
    expect(part(dialog, 'apply').classList.contains('acme-ok')).toBe(true);
    expect(dialog.querySelector('[data-probe]') !== null).toBe(true);
    type(field(dialog, 'count'), '2');
    await nextTick();
    expect(read as string).toBe('2');
  });
});

describe('the popups registry renders the packaged dialogs', () => {
  const keys = ['paragraph', 'font', 'bulletsAndNumbering', 'columns', 'background'] as const;
  const Opener = defineComponent({
    props: { dialog: { type: String as () => PptxEditorDialogKey, required: true } },
    setup(props) {
      const handle = useEditorDialog(props.dialog);
      return () =>
        h('button', {
          'data-testid': 'opener',
          'data-available': String(handle.isAvailable.value),
          onClick: () => handle.open(),
        });
    },
  });
  const opener = () => document.querySelector<HTMLElement>('[data-testid="opener"]')!;
  async function root(dialog: PptxEditorDialogKey, popups?: PptxEditorPopups): Promise<void> {
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, popups ? { popups } : {}, () => [h(Opener, { dialog })]),
      }) as Component
    );
    app.mount(host);
    await nextTick();
  }
  const reset = (): void => {
    app?.unmount();
    host?.remove();
    app = null;
    host = null;
  };

  for (const key of keys)
    test(`${key}: packaged by default, replaced by a renderer, off with false`, async () => {
      await root(key);
      opener().click();
      await nextTick();
      expect(document.querySelector(`dialog[data-pptx-dialog="${key}"]`) !== null).toBe(true);
      // Cancel closes it.
      document
        .querySelector<HTMLElement>(`[data-pptx-dialog="${key}"] [data-pptx-part="cancel"]`)!
        .click();
      await nextTick();
      expect(document.querySelector(`[data-pptx-dialog="${key}"]`) === null).toBe(true);
      reset();

      const Mine = defineComponent({
        props: { open: Boolean, onClose: Function },
        setup: () => () => h('div', { 'data-testid': 'mine' }),
      });
      await root(key, { [key]: definePopup(Mine) } as PptxEditorPopups);
      opener().click();
      await nextTick();
      expect(document.querySelector('[data-testid="mine"]') !== null).toBe(true);
      expect(document.querySelector(`[data-pptx-dialog="${key}"]`) === null).toBe(true);
      reset();

      await root(key, { [key]: false } as PptxEditorPopups);
      expect(opener().getAttribute('data-available')).toBe('false');
      opener().click();
      await nextTick();
      expect(document.querySelector('[data-pptx-dialog]') === null).toBe(true);
    });
});
