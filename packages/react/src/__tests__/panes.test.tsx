import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import type { ReactNode } from 'react';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditor, useEditorColorMode, useEditorDialog, useNotesPane } from '../index.ts';

// The pane props (`open`, `onOpenChange`, `paneWidth`) and `colorMode`.

afterEach(() => cleanup());

const deck = () => buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] });

async function mount(ui: (ready: (e: PptxEditorInstance) => void) => ReactNode) {
  let editor: PptxEditorInstance | null = null;
  render(<>{ui((e) => (editor = e))}</>);
  await act(async () => void (await editor!.load(deck())));
  return editor!;
}

describe('pane props', () => {
  test('the slide rail hides to a strip and comes back, and reports the change', async () => {
    const changes: boolean[] = [];
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail paneWidth={240} onOpenChange={(o) => changes.push(o)} />
      </PptxEditor.Root>
    ));
    const rail = document.querySelector<HTMLElement>('.pptx-rail')!;
    expect(rail.style.width).toBe('240px');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Hide thumbnails' })));
    expect(document.querySelector('.pptx-rail--collapsed') !== null).toBe(true);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Show thumbnails' })));
    expect(document.querySelector('.pptx-rail--collapsed') === null).toBe(true);
    expect(changes).toEqual([false, true]);
  });

  test('a horizontal slide rail steps with Left and Right and labels its moves', async () => {
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail orientation="horizontal" />
      </PptxEditor.Root>
    ));
    const rail = document.querySelector<HTMLElement>('.pptx-rail')!;
    expect(rail.dataset.orientation).toBe('horizontal');
    const list = rail.querySelector<HTMLElement>('.pptx-rail__list')!;
    expect(list.getAttribute('aria-orientation')).toBe('horizontal');
    expect(screen.getByRole('button', { name: 'Move slide left' }) !== null).toBe(true);
    expect(screen.getByRole('button', { name: 'Move slide right' }) !== null).toBe(true);
    const first = rail.querySelector<HTMLElement>('.pptx-rail__item')!;
    act(() => void fireEvent.keyDown(first, { key: 'ArrowRight' }));
    expect(editor.snapshot().currentSlide).toBe(1);
    act(() => void fireEvent.keyDown(first, { key: 'ArrowDown' }));
    expect(editor.snapshot().currentSlide).toBe(1);
    act(() => void fireEvent.keyDown(first, { key: 'ArrowLeft' }));
    expect(editor.snapshot().currentSlide).toBe(0);
    // Hide thumbnails folds the row down; the strip keeps the orientation.
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Hide thumbnails' })));
    const strip = document.querySelector<HTMLElement>('.pptx-rail--collapsed')!;
    expect(strip.dataset.orientation).toBe('horizontal');
  });

  test('a vertical slide rail keeps Up and Down', async () => {
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail />
      </PptxEditor.Root>
    ));
    const rail = document.querySelector<HTMLElement>('.pptx-rail')!;
    expect(rail.dataset.orientation).toBe('vertical');
    const first = rail.querySelector<HTMLElement>('.pptx-rail__item')!;
    act(() => void fireEvent.keyDown(first, { key: 'ArrowRight' }));
    expect(editor.snapshot().currentSlide).toBe(0);
    act(() => void fireEvent.keyDown(first, { key: 'ArrowDown' }));
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(screen.getByRole('button', { name: 'Move slide up' }) !== null).toBe(true);
  });

  test('a controlled notes pane follows the host', async () => {
    function Host({ onReady }: { onReady: (e: PptxEditorInstance) => void }) {
      const [open, setOpen] = useState(false);
      return (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.NotesPane open={open} onOpenChange={setOpen} className="acme-notes" />
          <span data-testid="open">{String(open)}</span>
        </PptxEditor.Root>
      );
    }
    await mount((onReady) => <Host onReady={onReady} />);
    expect(document.querySelector('.pptx-notes.acme-notes .pptx-notes__text') === null).toBe(true);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Show speaker notes' })));
    expect(screen.getByTestId('open').textContent).toBe('true');
    expect(document.querySelector('.pptx-notes__text') !== null).toBe(true);
  });
});

describe('colorMode', () => {
  test('dark puts dark beside the scope on the viewport and on a dialog outside it', async () => {
    function Opener() {
      const pageSetup = useEditorDialog('pageSetup');
      return <button type="button" data-testid="open" onClick={() => pageSetup.open()} />;
    }
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady} colorMode="dark">
        <PptxEditor.Viewport>
          <span />
        </PptxEditor.Viewport>
        <Opener />
      </PptxEditor.Root>
    ));
    expect(document.querySelector('.pptx-editor.dark') !== null).toBe(true);
    act(() => void fireEvent.click(screen.getByTestId('open')));
    const dialog = document.querySelector('[data-pptx-dialog="pageSetup"]')!;
    expect(dialog.classList.contains('pptx-editor')).toBe(true);
    expect(dialog.classList.contains('dark')).toBe(true);
  });

  test('light is the default', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.Viewport />
      </PptxEditor.Root>
    ));
    expect(document.querySelector('.pptx-editor')!.classList.contains('dark')).toBe(false);
  });
});

describe('pane props, review round 1', () => {
  test('closing the notes through useNotesPane writes the typed draft first', async () => {
    let handle: ReturnType<typeof useNotesPane> | null = null;
    function Grab() {
      handle = useNotesPane();
      return null;
    }
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.NotesPane />
        <Grab />
      </PptxEditor.Root>
    ));
    // The notes field is contenteditable: the browser edits the DOM, then fires `input`.
    const field = document.querySelector<HTMLElement>('.pptx-notes__text')!;
    act(() => field.focus());
    act(() => {
      field.textContent = 'typed notes';
      void fireEvent.input(field);
    });
    act(() => handle!.setOpen(false));
    expect(editor.snapshot().notes).toBe('typed notes');
  });

  test('a notes pane outside a root still collapses with its own state', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck());
    render(
      <PptxEditorContext.Provider value={editor}>
        <PptxEditor.NotesPane />
      </PptxEditorContext.Provider>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Hide speaker notes' })));
    expect(document.querySelector('.pptx-notes__text') === null).toBe(true);
  });

  test('a controlled notes pane and useNotesPane agree, and the hook reaches onOpenChange', async () => {
    let handle: ReturnType<typeof useNotesPane> | null = null;
    function Host({ onReady }: { onReady: (e: PptxEditorInstance) => void }) {
      const [open, setOpen] = useState(false);
      return (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.NotesPane open={open} onOpenChange={setOpen} />
          <Inner />
        </PptxEditor.Root>
      );
    }
    function Inner() {
      handle = useNotesPane();
      return null;
    }
    await mount((onReady) => <Host onReady={onReady} />);
    expect(handle!.open).toBe(false);
    act(() => handle!.toggle());
    expect(handle!.open).toBe(true);
    expect(document.querySelector('.pptx-notes__text') !== null).toBe(true);
  });

  test('defaultOpen={false} renders closed on the first render', async () => {
    const seen: boolean[] = [];
    function Spy() {
      seen.push(document.querySelector('.pptx-notes__text') !== null);
      return null;
    }
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.NotesPane defaultOpen={false} />
        <Spy />
      </PptxEditor.Root>
    ));
    expect(seen.every((shown) => !shown)).toBe(true);
    expect(document.querySelector('.pptx-notes__text') === null).toBe(true);
  });

  test('the rail toggle keeps focus on a toggle, and toggle={false} removes both', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail />
      </PptxEditor.Root>
    ));
    const hide = screen.getByRole('button', { name: 'Hide thumbnails' });
    hide.focus();
    act(() => void fireEvent.click(hide));
    expect(document.activeElement).toBe(screen.getByRole('button', { name: 'Show thumbnails' }));
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Show thumbnails' })));
    expect(document.activeElement).toBe(screen.getByRole('button', { name: 'Hide thumbnails' }));
    cleanup();
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail toggle={false} />
      </PptxEditor.Root>
    ));
    expect(screen.queryByRole('button', { name: 'Hide thumbnails' }) === null).toBe(true);
  });

  test('the side panel takes className and paneWidth', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SidePanel className="acme-panel" paneWidth={320} />
      </PptxEditor.Root>
    ));
    const panel = document.querySelector<HTMLElement>('.acme-panel')!;
    expect(panel.style.width).toBe('320px');
  });

  test('colorMode system follows the operating system, and useEditorColorMode reports it', async () => {
    const listeners = new Set<() => void>();
    let dark = true;
    const original = globalThis.matchMedia;
    globalThis.matchMedia = ((query: string) => ({
      matches: dark,
      media: query,
      addEventListener: (_: string, l: () => void) => listeners.add(l),
      removeEventListener: (_: string, l: () => void) => listeners.delete(l),
    })) as unknown as typeof matchMedia;
    let mode = '';
    function Probe() {
      mode = useEditorColorMode().colorMode;
      return null;
    }
    try {
      await mount((onReady) => (
        <PptxEditor.Root onReady={onReady} colorMode="system">
          <PptxEditor.Viewport>
            <Probe />
          </PptxEditor.Viewport>
        </PptxEditor.Root>
      ));
      expect(mode).toBe('dark');
      expect(document.querySelector('.pptx-editor.dark') !== null).toBe(true);
      dark = false;
      act(() => listeners.forEach((l) => l()));
      expect(mode).toBe('light');
      expect(document.querySelector('.pptx-editor.dark') === null).toBe(true);
      cleanup();
      expect(listeners.size).toBe(0);
    } finally {
      globalThis.matchMedia = original;
    }
  });

  test('useNotesPane outside a root warns', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      function Probe() {
        useNotesPane();
        return null;
      }
      render(<Probe />);
      expect(warn.mock.calls.map((c: unknown[]) => String(c[0])).join('')).toContain(
        'useNotesPane is used outside PptxEditor.Root'
      );
    } finally {
      warn.mockRestore();
    }
  });

  test('two controlled notes panes in one root render without a loop, and warn', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await mount((onReady) => (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.NotesPane open />
          <PptxEditor.NotesPane open={false} />
        </PptxEditor.Root>
      ));
      expect(document.querySelectorAll('.pptx-notes__text').length).toBe(1);
      expect(warn.mock.calls.map((c: unknown[]) => String(c[0])).join('')).toContain(
        'PptxEditor.NotesPane is rendered twice'
      );
    } finally {
      warn.mockRestore();
    }
  });

  test('a vetoed rail toggle does not take focus when the host closes the rail later', async () => {
    let close: (() => void) | null = null;
    function Host({ onReady }: { onReady: (e: PptxEditorInstance) => void }) {
      const [open, setOpen] = useState(true);
      close = () => setOpen(false);
      return (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.SlideRail open={open} onOpenChange={() => {}} />
          <button type="button" data-testid="host">
            host
          </button>
        </PptxEditor.Root>
      );
    }
    await mount((onReady) => <Host onReady={onReady} />);
    const hide = screen.getByRole('button', { name: 'Hide thumbnails' });
    hide.focus();
    act(() => void fireEvent.click(hide));
    expect(screen.getByRole('button', { name: 'Hide thumbnails' })).toBeTruthy();
    const host = screen.getByTestId('host');
    host.focus();
    act(() => close!());
    expect(document.activeElement).toBe(host);
  });

  test('panes write data-open as DOCX Editor does, and the zoom pill takes className', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail />
        <PptxEditor.NotesPane />
        <PptxEditor.ZoomPill className="acme-zoom" />
      </PptxEditor.Root>
    ));
    const rail = () => document.querySelector('.pptx-rail')!;
    expect(rail().getAttribute('data-open')).toBe('true');
    expect(document.querySelector('.pptx-notes')!.getAttribute('data-open')).toBe('true');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Hide thumbnails' })));
    expect(rail().getAttribute('data-open')).toBe('false');
    expect(document.querySelector('.acme-zoom') !== null).toBe(true);
    // The zoom menu opens on the current level, the checked row.
    act(() => void fireEvent.click(document.querySelector('.pptx-zoom__level')!));
    const checked = document.querySelector('[role="menuitemradio"][aria-checked="true"]');
    expect(checked !== null && document.activeElement === checked).toBe(true);
  });

  test('the toggles say whether their region shows, and point at it', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail />
        <PptxEditor.NotesPane />
      </PptxEditor.Root>
    ));
    const hide = screen.getByRole('button', { name: 'Hide thumbnails' });
    expect(hide.getAttribute('aria-expanded')).toBe('true');
    expect(document.getElementById(hide.getAttribute('aria-controls')!)!.className).toContain(
      'pptx-rail__list'
    );
    const notes = screen.getByRole('button', { name: 'Hide speaker notes' });
    expect(notes.getAttribute('aria-expanded')).toBe('true');
    act(() => void fireEvent.click(notes));
    expect(
      screen.getByRole('button', { name: 'Show speaker notes' }).getAttribute('aria-expanded')
    ).toBe('false');
  });

  test('the rail: defaultOpen, a scroll to the current slide on open, and no menu when closed', async () => {
    let setOpen: ((open: boolean) => void) | null = null;
    function Host({ onReady }: { onReady: (e: PptxEditorInstance) => void }) {
      const [open, set] = useState(false);
      setOpen = set;
      return (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.SlideRail open={open} onOpenChange={set} />
        </PptxEditor.Root>
      );
    }
    const editor = await mount((onReady) => <Host onReady={onReady} />);
    expect(document.querySelector('.pptx-rail')!.getAttribute('data-open')).toBe('false');
    act(() => editor.goToSlide(2));
    // A list 100px tall that scrolls, with slides 100px apart: slide 3 sits below the fold.
    // The rail scrolls its own list to it, never the page (`scrollIntoNearestScroller`).
    const style = document.createElement('style');
    style.textContent = '.pptx-rail__list { overflow-y: auto; }';
    document.head.append(style);
    // Patch own properties of HTMLElement.prototype and restore the exact descriptors, so no
    // stub or shadowing property stays behind for later test files (they patch Element.prototype).
    const proto = HTMLElement.prototype;
    const keys = ['getBoundingClientRect', 'scrollBy', 'scrollIntoView', 'scrollHeight'] as const;
    const saved = keys.map((key) => [key, Object.getOwnPropertyDescriptor(proto, key)] as const);
    const scrolled: { el: string; top: number }[] = [];
    const stub = (key: string, descriptor: PropertyDescriptor) =>
      Object.defineProperty(proto, key, { configurable: true, ...descriptor });
    stub('getBoundingClientRect', {
      value(this: HTMLElement) {
        const i = this.dataset.index;
        const top = i === undefined ? 0 : Number(i) * 100;
        return { top, left: 0, width: 100, height: i === undefined ? 100 : 90 } as DOMRect;
      },
    });
    stub('scrollBy', {
      value(this: HTMLElement, to?: ScrollToOptions | number) {
        scrolled.push({ el: this.className, top: typeof to === 'object' ? (to.top ?? 0) : 0 });
      },
    });
    stub('scrollIntoView', {
      value() {
        throw new Error('the rail must not scroll the page');
      },
    });
    stub('scrollHeight', {
      get(this: HTMLElement) {
        return this.classList.contains('pptx-rail__list') ? 300 : 0;
      },
    });
    try {
      act(() => setOpen!(true));
    } finally {
      for (const [key, descriptor] of saved) {
        if (descriptor) Object.defineProperty(proto, key, descriptor);
        else delete (proto as unknown as Record<string, unknown>)[key];
      }
      style.remove();
    }
    expect(scrolled.some((s) => s.el.includes('pptx-rail__list') && s.top > 0)).toBe(true);
    const item = document.querySelector<HTMLElement>('.pptx-rail__item')!;
    act(() => void fireEvent.contextMenu(item, { clientX: 20, clientY: 20 }));
    expect(document.querySelector('.pptx-menu') !== null).toBe(true);
    act(() => setOpen!(false));
    act(() => setOpen!(true));
    expect(document.querySelector('.pptx-menu') === null).toBe(true);
    cleanup();
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail defaultOpen={false} />
      </PptxEditor.Root>
    ));
    expect(document.querySelector('.pptx-rail')!.getAttribute('data-open')).toBe('false');
  });

  test('round 3: a closed rail with no toggle renders nothing', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail open={false} toggle={false} />
      </PptxEditor.Root>
    ));
    expect(document.querySelector('nav.pptx-rail') === null).toBe(true);
  });

  test('round 3: a refused hook toggle asks for the same value again', async () => {
    const asks: boolean[] = [];
    let handle: ReturnType<typeof useNotesPane> | null = null;
    function Grab() {
      handle = useNotesPane();
      return null;
    }
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.NotesPane open onOpenChange={(o) => asks.push(o)} />
        <Grab />
      </PptxEditor.Root>
    ));
    act(() => handle!.toggle());
    act(() => handle!.toggle());
    expect(asks).toEqual([false, false]);
  });

  test('round 3: a second uncontrolled notes pane keeps its own state', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await mount((onReady) => (
        <PptxEditor.Root onReady={onReady}>
          <PptxEditor.NotesPane className="first" />
          <PptxEditor.NotesPane className="second" />
        </PptxEditor.Root>
      ));
      const hide = document.querySelector<HTMLElement>(
        '.second .pptx-notes__head button[aria-expanded]'
      )!;
      act(() => void fireEvent.click(hide));
      expect(document.querySelector('.second')!.getAttribute('data-open')).toBe('false');
      expect(document.querySelector('.first')!.getAttribute('data-open')).toBe('true');
    } finally {
      warn.mockRestore();
    }
  });

  test('round 4: style on the panes, a width in style wins, and the strip stays narrow', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.SlideRail paneWidth={240} style={{ width: 300, top: 4 }} />
        <PptxEditor.NotesPane style={{ marginTop: 6 }} />
        <PptxEditor.SidePanel paneWidth={280} style={{ width: 320 }} />
      </PptxEditor.Root>
    ));
    const rail = () => document.querySelector<HTMLElement>('.pptx-rail')!;
    expect(rail().style.width).toBe('300px');
    expect(rail().style.top).toBe('4px');
    expect(document.querySelector<HTMLElement>('.pptx-notes')!.style.marginTop).toBe('6px');
    expect(document.querySelector<HTMLElement>('.pptx-panel')!.style.width).toBe('320px');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Hide thumbnails' })));
    expect(rail().style.width).toBe('');
    expect(rail().style.top).toBe('4px');
  });

  test('round 4: when the owning notes pane unmounts, the other one follows the hook', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    let handle: ReturnType<typeof useNotesPane> | null = null;
    let showFirst: ((show: boolean) => void) | null = null;
    function Host({ onReady }: { onReady: (e: PptxEditorInstance) => void }) {
      const [first, setFirst] = useState(true);
      showFirst = setFirst;
      return (
        <PptxEditor.Root onReady={onReady}>
          {first ? <PptxEditor.NotesPane className="first" /> : null}
          <PptxEditor.NotesPane className="second" />
          <Grab />
        </PptxEditor.Root>
      );
    }
    function Grab() {
      handle = useNotesPane();
      return null;
    }
    try {
      await mount((onReady) => <Host onReady={onReady} />);
      act(() => showFirst!(false));
      act(() => handle!.toggle());
      expect(document.querySelector('.second')!.getAttribute('data-open')).toBe(
        String(handle!.open)
      );
    } finally {
      warn.mockRestore();
    }
  });

  test('useNotesPane functions keep their identity when the notes open and close', async () => {
    const seen: Array<ReturnType<typeof useNotesPane>> = [];
    function Grab() {
      seen.push(useNotesPane());
      return null;
    }
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.NotesPane />
        <Grab />
      </PptxEditor.Root>
    ));
    const first = seen[seen.length - 1]!;
    act(() => first.toggle());
    const last = seen[seen.length - 1]!;
    expect(last.open).toBe(false);
    expect(last.setOpen).toBe(first.setOpen);
    expect(last.toggle).toBe(first.toggle);
  });
});
