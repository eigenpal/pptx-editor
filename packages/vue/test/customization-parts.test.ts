// The customization API with DOCX Editor's syntax, in Vue: the sugar's part props,
// `useHyperlinkPopup`, the named menu rows, the `t` of the find bar and the notices,
// `usePptxSource`, and `useEditorHostDocument`. The twin of the React test
// (packages/react/src/__tests__/customization-parts.test.tsx).

import { afterEach, describe, expect, test } from 'bun:test';
import {
  createApp,
  defineComponent,
  effectScope,
  h,
  nextTick,
  shallowRef,
  watch,
  type App,
} from 'vue';
import {
  defineFontResolver,
  requestLinkPopup,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import {
  PptxEditor,
  PptxEditorContextMenu,
  PptxEditorMenu,
  PptxEditorToolbar,
  requestFindBar,
  useEditorHostDocument,
  useHyperlinkPopup,
  useOpenFailure,
  usePptxSource,
  type ChromeTranslate,
} from '../src/index.ts';

const IN = 914_400;
const PX = 9_525;

const apps: [App, HTMLElement][] = [];
afterEach(() => {
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
  document.body.replaceChildren();
});

function mount(render: () => ReturnType<typeof h>): void {
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({ render });
  app.mount(host);
  apps.push([app, host]);
}

const settle = async () => {
  await nextTick();
  await nextTick();
};

const shapeDeck = () =>
  buildPptx({ slides: [{ shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN }) }] });
const textDeck = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
    ],
  });

/** Mount the sugar, open a deck, and give the slide stage a size for pointer events. */
async function sugar(props: Record<string, unknown>, bytes = shapeDeck()) {
  let editor: PptxEditorInstance | null = null;
  mount(() => h(PptxEditor, { ...props, onReady: (e: PptxEditorInstance) => (editor = e) }));
  await settle();
  expect((await editor!.load(bytes)).ok).toBe(true);
  await settle();
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement | null;
  if (stage)
    stage.getBoundingClientRect = () =>
      ({
        left: 0,
        top: 0,
        right: 1280,
        bottom: 720,
        width: 1280,
        height: 720,
        x: 0,
        y: 0,
      }) as DOMRect;
  return { editor: editor!, stage };
}

async function rightClick(stage: HTMLElement, x: number, y: number): Promise<void> {
  stage.dispatchEvent(
    new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: (x * IN) / PX,
      clientY: (y * IN) / PX,
    })
  );
  await settle();
}

const slots = (scope: ParentNode) =>
  [...scope.querySelectorAll('[data-slot]')].map((r) => r.getAttribute('data-slot'));

function selectText(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({
    kind: 'text',
    nodeId,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 2 },
  });
}

describe('the sugar part props (Vue)', () => {
  test('true or an object adds a part, and false removes it', async () => {
    await sugar({
      menu: true,
      toolbar: { className: 'acme-bar' },
      slideRail: false,
      notesPane: false,
    });
    expect(document.querySelector('.pptx-menubar') !== null).toBe(true);
    expect(document.querySelector('.acme-bar') !== null).toBe(true);
    expect(document.querySelector('.pptx-rail') === null).toBe(true);
    expect(document.querySelector('.pptx-notes') === null).toBe(true);
  });

  test('the defaults: a toolbar and a rail, and no menu bar', async () => {
    await sugar({});
    expect(document.querySelector('.pptx-menubar') === null).toBe(true);
    expect(document.querySelector('.pptx-default__dock') !== null).toBe(true);
    expect(document.querySelector('.pptx-rail') !== null).toBe(true);
  });

  test('contextMenu is the popups.contextMenu entry', async () => {
    const { stage } = await sugar({ contextMenu: { className: 'acme-menu' } });
    await rightClick(stage!, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu.acme-menu') !== null).toBe(true);
  });

  test('contextMenu false shows no menu', async () => {
    const off = await sugar({ contextMenu: false });
    await rightClick(off.stage!, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
  });

  test('a popups.contextMenu entry wins over the prop', async () => {
    const both = await sugar({
      contextMenu: false,
      popups: {
        contextMenu: (props: Record<string, unknown>) =>
          h(PptxEditorContextMenu, { ...props, className: 'wins' }),
      },
    });
    await rightClick(both.stage!, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu.wins') !== null).toBe(true);
  });

  test('hyperlinkPopup false removes the packaged popup', async () => {
    const { editor } = await sugar({ hyperlinkPopup: false }, textDeck());
    selectText(editor);
    requestLinkPopup(editor);
    await settle();
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
  });
});

const LinkPanel = defineComponent({
  setup() {
    const popup = useHyperlinkPopup();
    return () =>
      popup.state.value.mode === 'closed'
        ? null
        : h('div', { 'data-testid': 'host-link' }, [
            h('output', { 'data-testid': 'draft' }, popup.state.value.draft),
            h('button', {
              'data-testid': 'save',
              onClick: () => {
                popup.setUrl('https://example.test/');
                popup.commitEdit();
              },
            }),
          ]);
  },
});

const Mode = defineComponent({
  setup() {
    const popup = useHyperlinkPopup();
    return () => h('output', { 'data-testid': 'mode' }, popup.state.value.mode);
  },
});

describe('useHyperlinkPopup (Vue)', () => {
  async function root(child: ReturnType<typeof defineComponent>, popups?: { hyperlink: false }) {
    let editor: PptxEditorInstance | null = null;
    mount(() =>
      h(
        PptxEditor.Root,
        { ...(popups ? { popups } : {}), onReady: (e: PptxEditorInstance) => (editor = e) },
        () => [
          h(PptxEditor.Content),
          h(PptxEditorToolbar, { preset: false }, () => [h(PptxEditorToolbar.Link)]),
          h(child),
        ]
      )
    );
    await settle();
    expect((await editor!.load(textDeck())).ok).toBe(true);
    await settle();
    selectText(editor!);
    await settle();
    return editor!;
  }

  test('the packaged popup and a host composable share one state', async () => {
    await root(Mode);
    const mode = () => document.querySelector('[data-testid="mode"]')!.textContent;
    expect(mode()).toBe('closed');
    document.querySelector<HTMLElement>('[data-slot="text.link"]')!.click();
    await settle();
    expect(document.querySelector('.pptx-link-popup') !== null).toBe(true);
    expect(mode()).toBe('editing');
  });

  test('with popups.hyperlink false, a host panel opens for the link control and applies', async () => {
    const editor = await root(LinkPanel, { hyperlink: false });
    const link = document.querySelector<HTMLElement>('[data-slot="text.link"]')!;
    expect(link.getAttribute('aria-disabled')).toBe(null);
    link.click();
    await settle();
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    document.querySelector<HTMLElement>('[data-testid="save"]')!.click();
    await settle();
    expect(document.querySelector('[data-testid="host-link"]') === null).toBe(true);
    expect(editor.snapshot().format?.link).toBe('https://example.test/');
  });
});

describe('useHyperlinkPopup open() (Vue)', () => {
  test('a javascript: target never reaches an anchor or window.open', async () => {
    let popup!: ReturnType<typeof useHyperlinkPopup>;
    const Opener = defineComponent({
      setup() {
        popup = useHyperlinkPopup();
        return () => null;
      },
    });
    let editor: PptxEditorInstance | null = null;
    mount(() =>
      h(PptxEditor.Root, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
        h(PptxEditor.Content),
        h(Opener),
      ])
    );
    await settle();
    expect((await editor!.load(textDeck())).ok).toBe(true);
    await settle();
    const opened: unknown[] = [];
    const openBefore = window.open;
    window.open = ((...args: unknown[]) => {
      opened.push(args);
      return null;
    }) as typeof window.open;
    try {
      // The controller opens for the editor selection, never for a caller-supplied target.
      selectText(editor!);
      await settle();
      popup.open({ url: 'javascript:alert(1)', href: 'javascript:alert(1)' });
      await settle();
      expect(document.querySelector('.pptx-link-popup') !== null).toBe(true);
      const hrefs = [...document.querySelectorAll('a[href]')].map((a) => a.getAttribute('href'));
      expect(hrefs.some((h) => h!.startsWith('javascript:'))).toBe(false);
      expect(popup.state.value.href).toBeNull();
      popup.openTarget();
      expect(opened).toEqual([]);
    } finally {
      window.open = openBefore;
    }
  });
});

describe('a refused link shows once (Vue)', () => {
  test('the open popup shows the refusal, and the edit notice does not repeat it', async () => {
    let popup!: ReturnType<typeof useHyperlinkPopup>;
    const Probe = defineComponent({
      setup() {
        popup = useHyperlinkPopup();
        return () => null;
      },
    });
    let editor: PptxEditorInstance | null = null;
    mount(() =>
      h(PptxEditor.Root, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
        h(PptxEditor.Content),
        h(PptxEditorToolbar, { preset: false }, () => [h(PptxEditorToolbar.Link)]),
        h(PptxEditor.EditNotice),
        h(Probe),
      ])
    );
    await settle();
    expect((await editor!.load(textDeck())).ok).toBe(true);
    await settle();
    selectText(editor!);
    await settle();
    document.querySelector<HTMLElement>('[data-slot="text.link"]')!.click();
    await settle();
    popup.setUrl('javascript:alert(1)');
    popup.commitEdit();
    await settle();
    expect(editor!.snapshot().lastRejection?.reason).toBe('errors.unsafeLink');
    expect(document.querySelector('.pptx-link__error') !== null).toBe(true);
    expect(document.querySelector('.pptx-edit-notice') === null).toBe(true);
    // Closing the popup does not bring the refusal back as a notice.
    popup.close();
    await settle();
    expect(document.querySelector('.pptx-edit-notice') === null).toBe(true);
    // Another refusal still shows.
    editor!.exec({ type: 'insertShape', preset: 'noSuchPreset' });
    await settle();
    expect(document.querySelector('.pptx-edit-notice') !== null).toBe(true);
  });
});

describe('named menu rows (Vue)', () => {
  test('ContextMenu.Copy changes its row in place; Menu.Cut hides its row', async () => {
    const { stage } = await sugar({
      menu: true,
      popups: {
        contextMenu: (props: Record<string, unknown>) =>
          h(PptxEditorContextMenu, props, () => [
            h(PptxEditorContextMenu.Copy, {
              label: 'Duplicate to clipboard',
              className: 'acme-copy',
            }),
            h(PptxEditorContextMenu.Cut, { hidden: true }),
          ]),
      },
    });
    await rightClick(stage!, 1.5, 1.5);
    const copy = document.querySelector<HTMLElement>('.pptx-contextmenu .acme-copy')!;
    expect(copy.textContent).toContain('Duplicate to clipboard');
    const order = slots(document.querySelector('.pptx-contextmenu')!);
    expect(order).not.toContain('edit.cut');
    expect(order.indexOf('edit.copy')).toBe(order.indexOf('edit.paste') - 1);
  });

  test('Menu.Cut inside Menu.Edit hides its row', async () => {
    mount(() =>
      h(PptxEditor.Root, null, () => [
        h(PptxEditorMenu, null, () => [
          h(PptxEditorMenu.Edit, null, () => [h(PptxEditorMenu.Cut, { hidden: true })]),
        ]),
      ])
    );
    await settle();
    const edit = [...document.querySelectorAll<HTMLElement>('.pptx-menubar__trigger')].find(
      (b) => b.textContent === 'Edit'
    )!;
    edit.click();
    await settle();
    const panel = document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
    expect(slots(panel)).not.toContain('edit.cut');
    expect(slots(panel)).toContain('edit.copy');
  });
});

describe('the t prop of the find bar and the notices (Vue)', () => {
  const host: ChromeTranslate = (key) => `host:${key}`;

  test('FindBar and OpenFailure take their labels from t', async () => {
    let editor: PptxEditorInstance | null = null;
    mount(() =>
      h(PptxEditor.Root, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
        h('div', { class: 'pptx-editor' }, [
          h(PptxEditor.Content),
          h(PptxEditor.FindBar, { t: host }),
          h(PptxEditor.OpenFailure, { t: host }),
          h(PptxEditor.EditNotice, { t: host }),
          h(PptxEditor.FontNotice, { t: host }),
        ]),
      ])
    );
    await settle();
    await editor!.load(textDeck());
    await settle();
    requestFindBar(editor, false);
    await settle();
    expect(document.querySelector('[role="search"]')!.getAttribute('aria-label')).toBe(
      'host:find.label'
    );
    await editor!.load(new Uint8Array([1, 2, 3]));
    await settle();
    expect(document.body.textContent).toContain('host:openFailure.title');
  });

  test('useOpenFailure takes t', async () => {
    let editor: PptxEditorInstance | null = null;
    const Toast = defineComponent({
      setup() {
        const { message } = useOpenFailure({ t: host });
        return () => h('output', { 'data-testid': 'toast' }, message.value ?? '');
      },
    });
    mount(() =>
      h(PptxEditor.Root, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [h(Toast)])
    );
    await settle();
    await editor!.load(new Uint8Array([1, 2, 3]));
    await settle();
    expect(document.querySelector('[data-testid="toast"]')!.textContent).toStartWith('host:');
  });
});

describe('usePptxSource (Vue)', () => {
  test('bytes pass through at once; eager fonts hold the deck back', async () => {
    const scope = effectScope();
    const bytes = new Uint8Array([1]);
    let release: (value: { supportedFamilies: string[] }) => void = () => undefined;
    const loader = () => new Promise<{ supportedFamilies: string[] }>((r) => (release = r));
    const [plain, eager] = scope.run(() => [
      usePptxSource(bytes),
      usePptxSource(bytes, { fonts: loader }),
    ])!;
    expect(plain!.document.value).toBe(bytes);
    expect(plain!.isLoading.value).toBe(false);
    expect(eager!.document.value).toBeUndefined();
    release({ supportedFamilies: ['Acme Sans'] });
    await new Promise((r) => setTimeout(r, 0));
    expect(eager!.document.value).toBe(bytes);
    expect((eager!.fonts.value as { supportedFamilies?: string[] }).supportedFamilies).toContain(
      'Acme Sans'
    );
    scope.stop();
  });

  test('an on-demand resolver releases the deck; a failed URL is the error', async () => {
    const fetchBefore = globalThis.fetch;
    globalThis.fetch = (async () =>
      new Response(null, { status: 404, statusText: 'Not Found' })) as unknown as typeof fetch;
    const scope = effectScope();
    try {
      const bytes = new Uint8Array([1]);
      const [demand, missing] = scope.run(() => [
        usePptxSource(bytes, { fonts: defineFontResolver(() => undefined) }),
        usePptxSource('/missing.pptx'),
      ])!;
      expect(demand!.document.value).toBe(bytes);
      expect(typeof demand!.fonts.value).toBe('function');
      await new Promise((r) => setTimeout(r, 0));
      expect(missing!.error.value?.message).toContain('404');
    } finally {
      scope.stop();
      globalThis.fetch = fetchBefore;
    }
  });
});

describe('usePptxSource keeps the deck (Vue)', () => {
  test('a getter with a new fonts identity at each read resolves the fonts once', async () => {
    const scope = effectScope();
    const bytes = new Uint8Array([1]);
    const tick = shallowRef(0);
    let calls = 0;
    const seen: (Uint8Array | undefined)[] = [];
    const result = scope.run(() =>
      usePptxSource(bytes, () => {
        void tick.value;
        return {
          fonts: () => {
            calls += 1;
            return { supportedFamilies: ['Acme Sans'] };
          },
        };
      })
    )!;
    await new Promise((r) => setTimeout(r, 0));
    expect(result.document.value).toBe(bytes);
    // Record every value that the root would see.
    scope.run(() => watch(result.document, (d) => seen.push(d), { flush: 'sync' }));
    tick.value += 1;
    await settle();
    tick.value += 1;
    await new Promise((r) => setTimeout(r, 0));
    expect(calls).toBe(1);
    expect(seen).toEqual([]);
    expect(result.document.value).toBe(bytes);
    scope.stop();
  });

  test('the old bytes stay while a new URL loads', async () => {
    const fetchBefore = globalThis.fetch;
    let releaseB: () => void = () => undefined;
    globalThis.fetch = (async (url: string) => {
      if (String(url).endsWith('b.pptx')) await new Promise<void>((r) => (releaseB = r));
      return new Response(new Uint8Array(String(url).endsWith('b.pptx') ? [2] : [1]));
    }) as unknown as typeof fetch;
    const scope = effectScope();
    try {
      const url = shallowRef('/a.pptx');
      const result = scope.run(() => usePptxSource(url))!;
      await new Promise((r) => setTimeout(r, 0));
      expect(result.document.value).toEqual(new Uint8Array([1]));
      url.value = '/b.pptx';
      await settle();
      expect(result.isLoading.value).toBe(true);
      expect(result.document.value).toEqual(new Uint8Array([1]));
      releaseB();
      await new Promise((r) => setTimeout(r, 0));
      expect(result.document.value).toEqual(new Uint8Array([2]));
      expect(result.isLoading.value).toBe(false);
    } finally {
      scope.stop();
      globalThis.fetch = fetchBefore;
    }
  });
});

describe('useEditorHostDocument (Vue)', () => {
  test('is the document the slide is mounted in', async () => {
    let seen: Document | null = null;
    const Probe = defineComponent({
      setup() {
        const doc = useEditorHostDocument();
        return () => {
          seen = doc.value;
          return null;
        };
      },
    });
    let editor: PptxEditorInstance | null = null;
    mount(() =>
      h(PptxEditor.Root, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
        h(PptxEditor.Content),
        h(Probe),
      ])
    );
    await settle();
    await editor!.load(shapeDeck());
    await settle();
    expect(seen === document).toBe(true);
  });
});
