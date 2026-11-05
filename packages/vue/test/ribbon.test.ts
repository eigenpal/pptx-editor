import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import {
  Text,
  createApp,
  defineComponent,
  h,
  nextTick,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  ribbonTabs,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import {
  PptxEditorRibbon,
  PptxEditorRoot,
  PptxEditorStatusBar,
  PptxEditorToolbar,
  PptxEditorViewport,
} from '../src/index.ts';
import { providePptxEditor, useEditorState } from '../src/composables.ts';
import { tooltipText } from './tooltip-text.ts';

// The Vue twin of the ribbon (packages/react/src/__tests__/ribbon.test.tsx).

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
  return editor;
}

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
});

async function mount(editor: PptxEditorInstance, render: () => unknown): Promise<HTMLElement> {
  host = document.createElement('div');
  document.body.append(host);
  const Child = defineComponent({ setup: () => render });
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => h(Child);
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

const tabNames = (el: HTMLElement) =>
  [...el.querySelectorAll('[role="tab"]')].map((t) => t.textContent);
const panelGroups = (el: HTMLElement) =>
  [...el.querySelectorAll('[role="tabpanel"] .pptx-toolbar > [data-group]')].map((g) =>
    g.getAttribute('data-group')
  );

describe('PptxEditorRibbon', () => {
  test('the tabs follow the reference PDF, and a click selects a tab', async () => {
    const editor = await open();
    const changes: string[] = [];
    const el = await mount(editor, () =>
      h(PptxEditorRibbon, { onTabChange: (t: string) => changes.push(t) })
    );
    expect(tabNames(el)).toEqual(['Home', 'Insert', 'Design', 'Slide Show', 'View']);
    expect(panelGroups(el).slice(0, 3)).toEqual(['edit', 'format', 'slide']);
    expect(el.querySelector('[data-slot="edit.paste"]') !== null).toBe(true);
    expect(el.querySelector('.pptx-ribbon__qat [data-slot="history.undo"]') !== null).toBe(true);
    el.querySelector<HTMLElement>('[data-tab="design"]')!.click();
    await nextTick();
    // Design holds the slide size and background slots, which open dialogs.
    expect(el.querySelector('[role="tabpanel"]')!.getAttribute('data-tab')).toBe('design');
    expect(changes).toEqual(['design']);
  });

  test('Shape Format shows while shapes are selected', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorRibbon));
    expect(tabNames(el)).not.toContain('Shape Format');
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    expect(tabNames(el)).toContain('Shape Format');
    expect(el.querySelector('[data-tab="shapeFormat"]')!.hasAttribute('data-contextual')).toBe(
      true
    );
  });

  test('a gone contextual tab gives way to Home for good and tells the host', async () => {
    const editor = await open();
    const changes: string[] = [];
    const el = await mount(editor, () =>
      h(PptxEditorRibbon, { onTabChange: (t: string) => changes.push(t) })
    );
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    el.querySelector<HTMLElement>('[data-tab="shapeFormat"]')!.click();
    await nextTick();
    editor.setSelection({ kind: 'none' });
    await nextTick();
    await nextTick();
    expect(changes).toEqual(['shapeFormat', 'home']);
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    expect(el.querySelector('[role="tab"][aria-selected="true"]')!.getAttribute('data-tab')).toBe(
      'home'
    );
  });

  test('the focus moves to Home when the focused contextual tab goes away', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorRibbon));
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    const format = el.querySelector<HTMLElement>('[data-tab="shapeFormat"]')!;
    format.click();
    format.focus();
    await nextTick();
    editor.setSelection({ kind: 'none' });
    await nextTick();
    await nextTick();
    await nextTick();
    expect(document.activeElement?.getAttribute('data-tab')).toBe('home');
  });

  test('misuse warns, as in React, and a contextual override with its context off adds no tab', async () => {
    const editor = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const el = await mount(editor, () =>
        h(PptxEditorRibbon, null, () => [
          h(PptxEditorRibbon.Tab, { id: 'Home' }),
          h(PptxEditorRibbon.Tab, { id: 'mine', label: 'Mine', after: 'nowhere' }),
          h(PptxEditorRibbon.Tab, { id: 'view', after: 'home' }),
          h(PptxEditorRibbon.Tab, { id: 'gone', hidden: true }),
          h(PptxEditorRibbon.Tab, { id: 'shapeFormat' }),
          h('span'),
        ])
      );
      const text = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(text).toContain('PptxEditorRibbon.Tab id="Home" is not a built-in tab');
      expect(text).toContain('after="nowhere", but no tab has that id');
      expect(text).toContain('id="view" is a built-in tab, so its after is ignored');
      expect(text).toContain('id="gone" is hidden, but no tab has that id');
      expect(text).toContain('draws only PptxEditorRibbon.Tab children');
      expect(el.querySelector('[data-tab="shapeFormat"]') === null).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });

  test('a text child warns, as a string child does in React; blank text does not', async () => {
    const editor = await open();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await mount(editor, () => h(PptxEditorRibbon, null, () => [h(Text, '  ')]));
      expect(warn).not.toHaveBeenCalled();
      app?.unmount();
      host?.remove();
      await mount(editor, () => h(PptxEditorRibbon, null, () => [h(Text, 'loose text')]));
      expect(String(warn.mock.calls[0]?.[0])).toContain('draws only PptxEditorRibbon.Tab children');
    } finally {
      warn.mockRestore();
    }
  });

  test('an insert opens the Format tab of the new object', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorRibbon));
    editor.exec({ type: 'insertTextBox' });
    await nextTick();
    await nextTick();
    await nextTick();
    expect(el.querySelector('[role="tab"][aria-selected="true"]')!.getAttribute('data-tab')).toBe(
      'shapeFormat'
    );
  });

  test('outside a root, a dialog button shows disabled with the reason', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorRibbon, { defaultTab: 'design' }));
    const size = el.querySelector<HTMLElement>('[data-slot="file.pageSetup"]')!;
    expect(size.getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(size)).toContain('inside the editor root');
  });

  test('under a root, Slide Size opens Page setup and Format Background is on (packaged dialog)', async () => {
    let editor: PptxEditorInstance | null = null;
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
            h(PptxEditorViewport, null, () => [h(PptxEditorRibbon, { defaultTab: 'design' })]),
          ]),
      }) as Component
    );
    app.mount(host);
    await nextTick();
    await editor!.load(buildPptx({ slides: [{ shapes: '' }] }));
    await nextTick();
    const size = host.querySelector<HTMLElement>('[data-slot="file.pageSetup"]')!;
    expect(size.getAttribute('aria-disabled')).toBeNull();
    size.click();
    await nextTick();
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') !== null).toBe(true);
    const background = host.querySelector<HTMLElement>('[data-slot="slide.background"]')!;
    // The Vue root packages the Background dialog, as React's does.
    expect(background.getAttribute('aria-disabled')).toBeNull();
  });

  test('a popups change at runtime turns a dialog button off and on at once', async () => {
    let editor: PptxEditorInstance | null = null;
    const popups = shallowRef<Record<string, false> | undefined>(undefined);
    host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(
            PptxEditorRoot,
            { popups: popups.value, onReady: (e: PptxEditorInstance) => (editor = e) },
            () => [
              h(PptxEditorViewport, null, () => [h(PptxEditorRibbon, { defaultTab: 'design' })]),
            ]
          ),
      }) as Component
    );
    app.mount(host);
    await nextTick();
    await editor!.load(buildPptx({ slides: [{ shapes: '' }] }));
    await nextTick();
    const size = (): HTMLElement =>
      host!.querySelector<HTMLElement>('[data-slot="file.pageSetup"]')!;
    expect(size().getAttribute('aria-disabled')).toBeNull();
    popups.value = { pageSetup: false };
    await nextTick();
    await nextTick();
    expect(size().getAttribute('aria-disabled')).toBe('true');
    expect(tooltipText(size())).toContain('turned off');
    popups.value = undefined;
    await nextTick();
    await nextTick();
    expect(size().getAttribute('aria-disabled')).toBeNull();
  });

  test('a Tab child changes a tab with toolbar children and adds a host tab', async () => {
    const editor = await open();
    const el = await mount(editor, () =>
      h(PptxEditorRibbon, null, () => [
        h(PptxEditorRibbon.Tab, { id: 'home' }, () => [
          h((PptxEditorToolbar as unknown as Record<string, Component>).Strike!, { hidden: true }),
        ]),
        h(PptxEditorRibbon.Tab, { id: 'view', hidden: true }),
        h(PptxEditorRibbon.Tab, { id: 'legal', label: 'Legal', after: 'home' }, () => [
          h(PptxEditorToolbar.Group, { id: 'legal', label: 'Legal' }, () => [
            h(PptxEditorToolbar.Action, { label: 'Ask legal' }),
          ]),
        ]),
      ])
    );
    expect(tabNames(el)).toEqual(['Home', 'Legal', 'Insert', 'Design', 'Slide Show']);
    expect(el.querySelector('[data-slot="text.strike"]') === null).toBe(true);
    el.querySelector<HTMLElement>('[data-tab="legal"]')!.click();
    await nextTick();
    expect(el.querySelector('[aria-label="Ask legal"]') !== null).toBe(true);
  });
});

describe('PptxEditorStatusBar', () => {
  test('slide x of y and zoom', async () => {
    const editor = await open();
    const el = await mount(editor, () => h(PptxEditorStatusBar));
    expect(el.querySelector('[data-part="slideNumber"]')!.textContent).toBe('Slide 1 of 1');
    editor.setZoom(1);
    await nextTick();
    el.querySelector<HTMLElement>('[aria-label="Zoom in"]')!.click();
    expect(editor.snapshot().zoom).toBe(1.1);
    // The zoom part is the zoom pill: its level opens the zoom menu, whose first row fits.
    el.querySelector<HTMLElement>('.pptx-zoom__level')!.click();
    await nextTick();
    document.querySelector<HTMLElement>('[role="menuitemradio"]')!.click();
    expect(editor.snapshot().fitZoom).toBe(true);
  });

  test('useEditorState with a selector that builds a new array updates once per state change', async () => {
    const editor = await open();
    let renders = 0;
    let tabs: { readonly value: ReturnType<typeof ribbonTabs> } | null = null;
    const HostRibbon = defineComponent({
      setup() {
        const slice = useEditorState((s) => ribbonTabs(s));
        tabs = slice;
        return () => {
          renders += 1;
          return h('p', { 'data-testid': 'tabs' }, slice.value.map((t) => t.tab.id).join(','));
        };
      },
    });
    await mount(editor, () => h(HostRibbon));
    const text = () => document.querySelector('[data-testid="tabs"]')!.textContent ?? '';
    expect(text().includes('home')).toBe(true);
    const first = renders;
    const before = tabs!.value;
    editor.setSelection(editor.snapshot().selection);
    await nextTick();
    expect(tabs!.value === before).toBe(true);
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    expect(text().includes('shapeFormat')).toBe(true);
    expect(renders - first).toBe(1);
  });
});
