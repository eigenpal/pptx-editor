// The customization API with DOCX Editor's syntax: the sugar's part props (`toolbar`, `menu`,
// `contextMenu`, `hyperlinkPopup`, `slideRail`, `notesPane`), `useHyperlinkPopup`, the named
// menu rows (`Menu.Cut`, `ContextMenu.Cut`), the `t` of the find bar and the notices,
// `usePptxSource`, and `useEditorHostDocument`. The Vue twin is
// packages/vue/test/customization-parts.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import type { ReactNode } from 'react';
import {
  act,
  cleanup,
  fireEvent,
  render,
  renderHook,
  screen,
  waitFor,
} from '@testing-library/react';
import { defineFontResolver, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { requestLinkPopup } from '../chrome-actions.tsx';
import { requestFindBar } from '../FindBar.tsx';
import {
  PptxEditor,
  Toolbar,
  useEditorHostDocument,
  useHyperlinkPopup,
  useOpenFailure,
  usePptxSource,
  type ChromeTranslate,
  type PptxEditorProps,
} from '../index.ts';

const IN = 914_400;
const PX = 9_525;

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

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
async function sugar(props: Omit<PptxEditorProps, 'onReady'>, bytes = shapeDeck()) {
  let editor: PptxEditorInstance | null = null;
  render(<PptxEditor {...props} onReady={(e) => (editor = e)} />);
  await act(async () => void (await editor!.load(bytes)));
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

const rightClick = (stage: HTMLElement, x: number, y: number) =>
  act(() => {
    fireEvent.contextMenu(stage, { clientX: (x * IN) / PX, clientY: (y * IN) / PX });
  });

const slots = (scope: ParentNode) =>
  [...scope.querySelectorAll('[data-slot]')].map((r) => r.getAttribute('data-slot'));

/** Select the text of the first text box. */
function selectText(editor: PptxEditorInstance): void {
  const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
  act(() =>
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 2 },
    })
  );
}

describe('the sugar part props', () => {
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
    cleanup();
    // The defaults: a toolbar, a rail, and notes, and no menu bar.
    await sugar({});
    expect(document.querySelector('.pptx-menubar') === null).toBe(true);
    expect(document.querySelector('.pptx-default__dock') !== null).toBe(true);
    expect(document.querySelector('.pptx-rail') !== null).toBe(true);
    cleanup();
    await sugar({ toolbar: false, slideRail: { className: 'acme-rail' } });
    expect(document.querySelector('.pptx-default__dock') === null).toBe(true);
    expect(document.querySelector('.pptx-rail.acme-rail') !== null).toBe(true);
  });

  test('contextMenu is the popups.contextMenu entry: an object changes rows, false removes it', async () => {
    const { stage } = await sugar({
      contextMenu: { className: 'acme-menu', children: <PptxEditor.ContextMenu.Cut hidden /> },
    });
    rightClick(stage!, 1.5, 1.5);
    const menu = document.querySelector<HTMLElement>('.pptx-contextmenu.acme-menu')!;
    expect(menu !== null).toBe(true);
    expect(slots(menu)).not.toContain('edit.cut');
    expect(slots(menu)).toContain('edit.copy');
    cleanup();
    const off = await sugar({ contextMenu: false });
    rightClick(off.stage!, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
    cleanup();
    // A popups entry wins over the prop.
    const both = await sugar({
      contextMenu: false,
      popups: { contextMenu: (props) => <PptxEditor.ContextMenu {...props} className="wins" /> },
    });
    rightClick(both.stage!, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu.wins') !== null).toBe(true);
  });

  test('hyperlinkPopup false removes the packaged popup', async () => {
    const { editor } = await sugar({ hyperlinkPopup: false }, textDeck());
    selectText(editor);
    act(() => requestLinkPopup(editor));
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
  });
});

function LinkPanel() {
  const { state, setUrl, commitEdit, close } = useHyperlinkPopup();
  if (state.mode === 'closed') return null;
  return (
    <form
      data-testid="host-link"
      onSubmit={(e) => {
        e.preventDefault();
        commitEdit();
      }}
    >
      <input aria-label="Address" value={state.draft} onChange={(e) => setUrl(e.target.value)} />
      <button type="submit">Save</button>
      <button type="button" onClick={close}>
        Cancel
      </button>
    </form>
  );
}

function Mode() {
  return <output data-testid="mode">{useHyperlinkPopup().state.mode}</output>;
}

describe('useHyperlinkPopup', () => {
  async function root(ui: ReactNode, popups?: { hyperlink: false }) {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditor.Root onReady={(e) => (editor = e)} {...(popups ? { popups } : {})}>
        <PptxEditor.Content />
        <Toolbar preset={false}>
          <Toolbar.Link />
        </Toolbar>
        {ui}
      </PptxEditor.Root>
    );
    await act(async () => void (await editor!.load(textDeck())));
    selectText(editor!);
    return editor!;
  }

  test('the packaged popup and a host hook share one state', async () => {
    await root(<Mode />);
    expect(screen.getByTestId('mode').textContent).toBe('closed');
    act(() => void fireEvent.click(document.querySelector('[data-slot="text.link"]')!));
    expect(document.querySelector('.pptx-link-popup') !== null).toBe(true);
    expect(screen.getByTestId('mode').textContent).toBe('editing');
  });

  test('with popups.hyperlink false, a host panel opens for the link control and applies', async () => {
    const editor = await root(<LinkPanel />, { hyperlink: false });
    const link = document.querySelector<HTMLElement>('[data-slot="text.link"]')!;
    // A mounted panel keeps the control on.
    expect(link.getAttribute('aria-disabled')).toBe(null);
    act(() => void fireEvent.click(link));
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    act(
      () =>
        void fireEvent.change(screen.getByLabelText('Address'), {
          target: { value: 'https://example.test/' },
        })
    );
    act(() => void fireEvent.click(screen.getByText('Save')));
    expect(screen.queryByTestId('host-link') === null).toBe(true);
    expect(editor.snapshot().format?.link).toBe('https://example.test/');
  });
});

describe('useHyperlinkPopup open()', () => {
  test('a javascript: target never reaches an anchor or window.open', async () => {
    let popup: ReturnType<typeof useHyperlinkPopup> | null = null;
    function Opener() {
      popup = useHyperlinkPopup();
      return null;
    }
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditor.Root onReady={(e) => (editor = e)}>
        <PptxEditor.Content />
        <Opener />
      </PptxEditor.Root>
    );
    await act(async () => void (await editor!.load(textDeck())));
    selectText(editor!);
    const opened: unknown[] = [];
    const openBefore = window.open;
    window.open = ((...args: unknown[]) => {
      opened.push(args);
      return null;
    }) as typeof window.open;
    try {
      act(() => popup!.open({ url: 'javascript:alert(1)', href: 'javascript:alert(1)' }));
      expect(document.querySelector('.pptx-link-popup') !== null).toBe(true);
      const hrefs = [...document.querySelectorAll('a[href]')].map((a) => a.getAttribute('href'));
      expect(hrefs.some((h) => h!.startsWith('javascript:'))).toBe(false);
      expect(popup!.state.href).toBeNull();
      act(() => void popup!.openTarget());
      expect(opened).toEqual([]);
    } finally {
      window.open = openBefore;
    }
  });
});

describe('a refused link shows once', () => {
  test('the open popup shows the refusal, and the edit notice does not repeat it', async () => {
    let popup: ReturnType<typeof useHyperlinkPopup> | null = null;
    function Probe() {
      popup = useHyperlinkPopup();
      return null;
    }
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditor.Root onReady={(e) => (editor = e)}>
        <PptxEditor.Content />
        <Toolbar preset={false}>
          <Toolbar.Link />
        </Toolbar>
        <PptxEditor.EditNotice />
        <Probe />
      </PptxEditor.Root>
    );
    await act(async () => void (await editor!.load(textDeck())));
    selectText(editor!);
    act(() => void fireEvent.click(document.querySelector('[data-slot="text.link"]')!));
    act(() => popup!.setUrl('javascript:alert(1)'));
    act(() => void popup!.commitEdit());
    expect(editor!.snapshot().lastRejection?.reason).toBe('errors.unsafeLink');
    expect(document.querySelector('.pptx-link__error') !== null).toBe(true);
    expect(document.querySelector('.pptx-edit-notice') === null).toBe(true);
    // Closing the popup does not bring the refusal back as a notice.
    act(() => popup!.close());
    expect(document.querySelector('.pptx-edit-notice') === null).toBe(true);
    // Another refusal still shows.
    act(() => void editor!.exec({ type: 'insertShape', preset: 'noSuchPreset' }));
    expect(document.querySelector('.pptx-edit-notice') !== null).toBe(true);
  });
});

describe('named menu rows', () => {
  test('ContextMenu.Copy changes its row in place; Menu.Cut hides its row', async () => {
    const { stage } = await sugar({
      menu: {
        children: (
          <PptxEditor.Menu.Edit>
            <PptxEditor.Menu.Cut hidden />
          </PptxEditor.Menu.Edit>
        ),
      },
      contextMenu: {
        children: (
          <PptxEditor.ContextMenu.Copy label="Duplicate to clipboard" className="acme-copy" />
        ),
      },
    });
    rightClick(stage!, 1.5, 1.5);
    const copy = document.querySelector<HTMLElement>('.pptx-contextmenu .acme-copy')!;
    expect(copy.textContent).toContain('Duplicate to clipboard');
    // In place: Copy stays between Cut and Paste.
    const order = slots(document.querySelector('.pptx-contextmenu')!);
    expect(order.indexOf('edit.copy')).toBe(order.indexOf('edit.cut') + 1);
    act(() => void fireEvent.click(document.querySelector('.pptx-contextmenu .acme-copy')!));
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Edit' })));
    const panel = document.querySelector<HTMLElement>('.pptx-menubar__panel')!;
    expect(slots(panel)).not.toContain('edit.cut');
    expect(slots(panel)).toContain('edit.copy');
  });

  test('a value row names its value: ContextMenu.BringToFront', async () => {
    const { stage } = await sugar({
      contextMenu: { children: <PptxEditor.ContextMenu.BringToFront hidden /> },
    });
    rightClick(stage!, 1.5, 1.5);
    const order = slots(document.querySelector('.pptx-contextmenu')!);
    expect(order).not.toContain('arrange.order.front');
    expect(order).toContain('arrange.order.back');
  });
});

describe('the t prop of the find bar and the notices', () => {
  const host: ChromeTranslate = (key) => `host:${key}`;

  test('FindBar and OpenFailure take their labels from t', async () => {
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditor.Root onReady={(e) => (editor = e)}>
        <div className="pptx-editor">
          <PptxEditor.Content />
          <PptxEditor.FindBar t={host} />
          <PptxEditor.OpenFailure t={host} />
          <PptxEditor.EditNotice t={host} />
          <PptxEditor.FontNotice t={host} />
        </div>
      </PptxEditor.Root>
    );
    await act(async () => void (await editor!.load(textDeck())));
    act(() => requestFindBar(editor, false));
    expect(screen.getByRole('search').getAttribute('aria-label')).toBe('host:find.label');
    await act(async () => void (await editor!.load(new Uint8Array([1, 2, 3]))));
    expect(document.body.textContent).toContain('host:openFailure.title');
  });

  test('useOpenFailure takes t', async () => {
    let editor: PptxEditorInstance | null = null;
    function Toast() {
      const { message } = useOpenFailure({ t: host });
      return <output data-testid="toast">{message}</output>;
    }
    render(
      <PptxEditor.Root onReady={(e) => (editor = e)}>
        <Toast />
      </PptxEditor.Root>
    );
    await act(async () => void (await editor!.load(new Uint8Array([1, 2, 3]))));
    expect(screen.getByTestId('toast').textContent).toStartWith('host:');
  });
});

describe('usePptxSource', () => {
  test('bytes pass through at once, with no fonts to wait for', () => {
    const bytes = new Uint8Array([1]);
    const { result } = renderHook(() => usePptxSource(bytes));
    expect(result.current.document).toBe(bytes);
    expect(result.current.isLoading).toBe(false);
    expect(result.current.fonts).toBeUndefined();
  });

  test('eager fonts hold the deck back until they settle', async () => {
    const bytes = new Uint8Array([1]);
    let release: (value: { supportedFamilies: string[] }) => void = () => undefined;
    const loader = () => new Promise<{ supportedFamilies: string[] }>((r) => (release = r));
    const { result } = renderHook(() => usePptxSource(bytes, { fonts: loader }));
    expect(result.current.document).toBeUndefined();
    expect(result.current.isLoading).toBe(true);
    await act(async () => release({ supportedFamilies: ['Acme Sans'] }));
    await waitFor(() => expect(result.current.document).toBe(bytes));
    expect((result.current.fonts as { supportedFamilies?: string[] }).supportedFamilies).toContain(
      'Acme Sans'
    );
  });

  test('an on-demand resolver releases the deck and keeps one identity', () => {
    const bytes = new Uint8Array([1]);
    const { result, rerender } = renderHook(() =>
      usePptxSource(bytes, { fonts: defineFontResolver(() => undefined) })
    );
    expect(result.current.document).toBe(bytes);
    const first = result.current.fonts;
    expect(typeof first).toBe('function');
    rerender();
    expect(result.current.fonts).toBe(first);
  });

  test('a URL is fetched, and a failed response is the error', async () => {
    const fetchBefore = globalThis.fetch;
    globalThis.fetch = (async (url: string) =>
      String(url).endsWith('missing.pptx')
        ? new Response(null, { status: 404, statusText: 'Not Found' })
        : new Response(new Uint8Array([7, 8]))) as unknown as typeof fetch;
    try {
      const ok = renderHook(() => usePptxSource('/deck.pptx'));
      await waitFor(() => expect(ok.result.current.document).toEqual(new Uint8Array([7, 8])));
      const missing = renderHook(() => usePptxSource('/missing.pptx'));
      await waitFor(() => expect(missing.result.current.error?.message).toContain('404'));
      expect(missing.result.current.document).toBeUndefined();
    } finally {
      globalThis.fetch = fetchBefore;
    }
  });
});

describe('usePptxSource keeps the deck', () => {
  test('a new fonts identity at each render resolves the fonts once', async () => {
    const bytes = new Uint8Array([1]);
    let calls = 0;
    const seen: (Uint8Array | undefined)[] = [];
    const { result, rerender } = renderHook(() => {
      const out = usePptxSource(bytes, {
        fonts: () => {
          calls += 1;
          return { supportedFamilies: ['Acme Sans'] };
        },
      });
      seen.push(out.document);
      return out;
    });
    await waitFor(() => expect(result.current.document).toBe(bytes));
    seen.length = 0;
    rerender();
    rerender();
    await act(async () => await new Promise((r) => setTimeout(r, 0)));
    expect(calls).toBe(1);
    expect(seen.every((d) => d === bytes)).toBe(true);
  });

  test('the old bytes stay while a new URL loads', async () => {
    const fetchBefore = globalThis.fetch;
    let releaseB: () => void = () => undefined;
    globalThis.fetch = (async (url: string) => {
      if (String(url).endsWith('b.pptx')) await new Promise<void>((r) => (releaseB = r));
      return new Response(new Uint8Array(String(url).endsWith('b.pptx') ? [2] : [1]));
    }) as unknown as typeof fetch;
    try {
      const { result, rerender } = renderHook(({ url }) => usePptxSource(url), {
        initialProps: { url: '/a.pptx' },
      });
      await waitFor(() => expect(result.current.document).toEqual(new Uint8Array([1])));
      rerender({ url: '/b.pptx' });
      expect(result.current.isLoading).toBe(true);
      expect(result.current.document).toEqual(new Uint8Array([1]));
      await act(async () => releaseB());
      await waitFor(() => expect(result.current.document).toEqual(new Uint8Array([2])));
      expect(result.current.isLoading).toBe(false);
    } finally {
      globalThis.fetch = fetchBefore;
    }
  });
});

describe('useEditorHostDocument', () => {
  test('is the document the slide is mounted in', async () => {
    let seen: Document | null = null;
    function Probe() {
      seen = useEditorHostDocument();
      return null;
    }
    let editor: PptxEditorInstance | null = null;
    render(
      <PptxEditor.Root onReady={(e) => (editor = e)}>
        <PptxEditor.Content />
        <Probe />
      </PptxEditor.Root>
    );
    await act(async () => void (await editor!.load(shapeDeck())));
    expect(seen === document).toBe(true);
  });
});
