// Every Vue part whose React twin takes `className` takes it too, and puts it on the same
// element as React: the element with the probe class has the same tag and the same other
// classes in both adapters. After DOCX Editor's Vue parts, where `className` merges with the
// host `class` on one element.

import { afterEach, expect, test } from 'bun:test';
import { createElement, type FunctionComponent, type ReactNode } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type VNode } from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { ContextMenuTargetProvider as ReactTargetProvider } from '../../react/src/context-menu-parts.tsx';
import * as R from '../../react/src/index.ts';
import * as V from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { ContextMenuTargetProvider as VueTargetProvider } from '../src/context-menu-parts.tsx';

const PROBE = 'probe';
const IN = 914_400;
const noop = (): void => undefined;

/** A React element of `part`, typed loosely: the cases pass each part its own props. */
const rc = (part: unknown, props: object | null, ...children: ReactNode[]): ReactNode =>
  createElement(
    part as FunctionComponent<Record<string, unknown>>,
    props as Record<string, unknown>,
    ...children
  );

let app: App | null = null;
let vueHost: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  vueHost?.remove();
  app = null;
  vueHost = null;
});

/** An editor with one text slide, whose snapshot also lists a substituted family. */
async function open(): Promise<PptxEditorInstance> {
  const real = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });
  expect((await real.load(bytes)).ok).toBe(true);
  const editor = Object.create(real) as PptxEditorInstance;
  let seen: unknown = null;
  let cached: ReturnType<PptxEditorInstance['snapshot']> | null = null;
  editor.snapshot = () => {
    const now = real.snapshot();
    if (now !== seen || !cached) {
      seen = now;
      cached = {
        ...now,
        substitutedFonts: [
          { family: 'Acme Display', substitute: 'Carlito', metricCompatible: false },
        ] as never,
      };
    }
    return cached;
  };
  return editor;
}

interface Case {
  readonly react: () => ReactNode;
  readonly vue: () => VNode;
  /** Runs after mount: opens what renders the class. */
  readonly after?: (editor: PptxEditorInstance) => void;
  /** A selector to click after mount. */
  readonly click?: string;
}

const p = { className: PROBE };
const menuState = { at: { x: 0, y: 0 }, keyboard: false };
const layouts = [{ id: 'probe', name: 'Probe', type: 'blank', masterId: 'm', masterName: 'M' }];

const CASES: Record<string, Case> = {
  Viewport: {
    react: () => rc(R.PptxEditor.Viewport, p),
    vue: () => h(V.PptxEditor.Viewport, p),
  },
  Content: {
    react: () => rc(R.PptxEditor.Content, p),
    vue: () => h(V.PptxEditor.Content, p),
  },
  Toolbar: {
    react: () => rc(R.PptxEditor.Toolbar, p),
    vue: () => h(V.PptxEditor.Toolbar, p),
  },
  'Toolbar preset false': {
    react: () => rc(R.PptxEditor.Toolbar, { ...p, preset: false }),
    vue: () => h(V.PptxEditor.Toolbar, { ...p, preset: false }),
  },
  'Toolbar.Separator': {
    react: () => rc(R.ToolbarSeparator, p),
    vue: () => h(V.ToolbarSeparator, p),
  },
  'Toolbar.Bold': {
    react: () => rc(R.PptxEditor.Toolbar.Bold, p),
    vue: () => h(V.PptxEditorToolbar.Bold!, p),
  },
  Ribbon: {
    react: () => rc(R.PptxEditor.Ribbon, p),
    vue: () => h(V.PptxEditor.Ribbon, p),
  },
  RibbonTab: {
    react: () => rc(R.PptxEditor.Ribbon, null, rc(R.RibbonTab, { id: 'home', ...p })),
    vue: () => h(V.PptxEditor.Ribbon, null, () => [h(V.RibbonTab, { id: 'home', ...p })]),
  },
  StatusBar: {
    react: () => rc(R.PptxEditor.StatusBar, p),
    vue: () => h(V.PptxEditor.StatusBar, p),
  },
  StatusBarSlideNumber: {
    react: () => rc(R.StatusBarSlideNumber, p),
    vue: () => h(V.StatusBarSlideNumber, p),
  },
  StatusBarNotes: {
    react: () => rc(R.StatusBarNotes, p),
    vue: () => h(V.StatusBarNotes, p),
  },
  StatusBarViews: {
    react: () => rc(R.StatusBarViews, p),
    vue: () => h(V.StatusBarViews, p),
  },
  StatusBarZoom: {
    react: () => rc(R.StatusBarZoom, p),
    vue: () => h(V.StatusBarZoom, p),
  },
  Menu: {
    react: () => rc(R.PptxEditor.Menu, p),
    vue: () => h(V.PptxEditor.Menu, p),
  },
  MenuEntry: {
    react: () =>
      createElement(
        R.PptxEditor.Menu,
        { preset: false },
        rc(R.MenuEntry, { id: 'acme', label: 'Acme', ...p })
      ),
    vue: () =>
      h(V.PptxEditor.Menu, { preset: false }, () => [
        h(V.MenuEntry, { id: 'acme', label: 'Acme', ...p }),
      ]),
    click: '[data-menu="acme"]',
  },
  MenuGroup: {
    react: () => rc(R.MenuGroup, { label: 'Group', ...p }),
    vue: () => h(V.MenuGroup, { label: 'Group', ...p }),
  },
  MenuSlotItem: {
    react: () => rc(R.MenuSlotItem, { slotId: 'text.bold', ...p }),
    vue: () => h(V.MenuSlotItem, { slotId: 'text.bold', ...p }),
  },
  ContextMenu: {
    react: () => rc(R.ContextMenu, { ...menuState, label: 'Menu', onClose: noop, ...p }),
    vue: () => h(V.ContextMenu, { ...menuState, label: 'Menu', onClose: noop, ...p }),
  },
  PptxEditorContextMenu: {
    react: () =>
      createElement(
        ReactTargetProvider,
        { value: { ...menuState, kind: 'slide', owner: document.body, close: noop } },
        rc(R.PptxEditor.ContextMenu, p)
      ),
    vue: () =>
      h(
        VueTargetProvider,
        { value: { ...menuState, kind: 'slide', owner: document.body, close: noop } },
        () => [h(V.PptxEditor.ContextMenu, p)]
      ),
  },
  SlideRail: {
    react: () => rc(R.PptxEditor.SlideRail, p),
    vue: () => h(V.PptxEditor.SlideRail, p),
  },
  'SlideRail closed': {
    react: () => rc(R.PptxEditor.SlideRail, { ...p, open: false }),
    vue: () => h(V.PptxEditor.SlideRail, { ...p, open: false }),
  },
  SlideSorter: {
    react: () => rc(R.PptxEditor.SlideSorter, p),
    vue: () => h(V.PptxEditor.SlideSorter, p),
  },
  NotesPane: {
    react: () => rc(R.PptxEditor.NotesPane, p),
    vue: () => h(V.PptxEditor.NotesPane, p),
  },
  SidePanel: {
    react: () => rc(R.PptxEditor.SidePanel, p),
    vue: () => h(V.PptxEditor.SidePanel, p),
  },
  ZoomPill: {
    react: () => rc(R.PptxEditor.ZoomPill, p),
    vue: () => h(V.PptxEditor.ZoomPill, p),
  },
  Loading: {
    react: () => rc(R.PptxEditor.Loading, { ...p, when: true }),
    vue: () => h(V.PptxEditor.Loading, { ...p, when: true }),
  },
  'Loading.Spinner': {
    react: () => rc(R.LoadingSpinner, p),
    vue: () => h(V.PptxEditorLoadingSpinner, p),
  },
  EditNotice: {
    react: () => rc(R.PptxEditor.EditNotice, p),
    vue: () => h(V.PptxEditor.EditNotice, p),
    after: (editor) => void editor.reject('disabled.selectionChanged'),
  },
  FontNotice: {
    react: () => rc(R.PptxEditor.FontNotice, p),
    vue: () => h(V.PptxEditor.FontNotice, p),
  },
  FindBar: {
    react: () => rc(R.PptxEditor.FindBar, p),
    vue: () => h(V.PptxEditor.FindBar, p),
  },
  HyperlinkPopup: {
    react: () => rc(R.HyperlinkPopup, { anchor: null, url: '', onClose: noop, ...p }),
    vue: () => h(V.HyperlinkPopup, { anchor: null, url: '', onClose: noop, ...p }),
  },
  LayoutGallery: {
    react: () => rc(R.LayoutGallery, { layouts, onChange: noop, ...p } as never),
    vue: () => h(V.LayoutGallery, { layouts, ...p } as never),
  },
  MenuTableGrid: {
    react: () => rc(R.MenuTableGrid, p),
    vue: () => h(V.MenuTableGrid, p),
  },
  ShapeStyleGallery: {
    react: () => rc(R.ShapeStyleGallery, { onPick: noop, ...p }),
    vue: () => h(V.ShapeStyleGallery, { onPick: noop, ...p }),
  },
};

for (const dialog of [
  'PageSetupDialog',
  'BackgroundDialog',
  'ImageAltTextDialog',
  'ImagePropertiesDialog',
  'ParagraphDialog',
  'FontDialog',
  'BulletsAndNumberingDialog',
  'ColumnsDialog',
] as const) {
  const props = { open: true, onClose: noop, ...p };
  CASES[dialog] = {
    react: () => rc(R.PptxEditor[dialog], props),
    vue: () => h(V.PptxEditor[dialog] as never, props),
  };
}

/** Each element that carries the probe class: its tag and its other classes. */
function landing(): string[] {
  return [...document.querySelectorAll(`.${PROBE}`)].map((el) =>
    [el.localName, ...[...el.classList].filter((c) => c !== PROBE && c !== 'dark').sort()].join('.')
  );
}

async function reactLanding(editor: PptxEditorInstance, c: Case): Promise<string[]> {
  const { container } = render(
    createElement(
      PptxEditorContext.Provider,
      { value: editor },
      createElement('div', { className: 'pptx-editor' }, c.react())
    )
  );
  if (c.after) act(() => c.after!(editor));
  if (c.click) act(() => container.querySelector<HTMLElement>(c.click!)?.click());
  const out = landing();
  cleanup();
  return out;
}

async function vueLanding(editor: PptxEditorInstance, c: Case): Promise<string[]> {
  vueHost = document.createElement('div');
  vueHost.className = 'pptx-editor';
  document.body.append(vueHost);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return () => c.vue();
      },
    })
  );
  app.mount(vueHost);
  await nextTick();
  if (c.after) {
    c.after(editor);
    await nextTick();
  }
  if (c.click) {
    vueHost.querySelector<HTMLElement>(c.click)?.click();
    await nextTick();
  }
  const out = landing();
  app.unmount();
  vueHost.remove();
  app = null;
  vueHost = null;
  return out;
}

test('className lands on the same element in React and Vue, for every part', async () => {
  const editor = await open();
  const differ: string[] = [];
  for (const [name, c] of Object.entries(CASES)) {
    const react = await reactLanding(editor, c);
    const vue = await vueLanding(editor, c);
    if (react.length !== 1 || JSON.stringify(react) !== JSON.stringify(vue))
      differ.push(`${name}: React ${JSON.stringify(react)}, Vue ${JSON.stringify(vue)}`);
  }
  expect(differ).toEqual([]);
});
