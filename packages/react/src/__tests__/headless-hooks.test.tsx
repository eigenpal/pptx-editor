// The hooks with DOCX Editor's names for a host's own chrome: `useChromeTranslate` and the
// compounds' `t`, the toolbar label hooks, `useHistoryGroup`, `useEditorCaret`, and
// `useDocumentSearch`. The Vue twin is packages/vue/test/headless-hooks.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { useState, type ReactNode } from 'react';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import {
  PptxEditor,
  useChromeTranslate,
  useDocumentSearch,
  useEditorCaret,
  useHistoryGroup,
  useToolbarContext,
  useToolbarLabel,
  type ChromeTranslate,
} from '../index.ts';

afterEach(() => cleanup());

const IN = 914400;
const textDeck = () =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('one two one')),
        }),
      },
      {
        shapes: shape({
          id: 2,
          name: 'U',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('one')),
        }),
      },
    ],
  });

async function mount(ui: (ready: (e: PptxEditorInstance) => void) => ReactNode) {
  let editor: PptxEditorInstance | null = null;
  render(<>{ui((e) => (editor = e))}</>);
  await act(async () => void (await editor!.load(textDeck())));
  return editor!;
}

const host: ChromeTranslate = (key) => `host:${key}`;

function Labels({ id }: { id: string }) {
  const t = useChromeTranslate();
  const label = useToolbarLabel();
  const context = useToolbarContext();
  return (
    <output data-testid={id}>
      {`${t('statusBar.notes')}|${label('statusBar.notes')}|${context.t ? 'host' : 'none'}`}
    </output>
  );
}

describe('the t prop and the translation hooks', () => {
  test('a compound with t labels itself and its parts from the host', async () => {
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.StatusBar t={host}>
          <Labels id="inside" />
        </PptxEditor.StatusBar>
        <PptxEditor.SlideRail t={host} />
        <Labels id="outside" />
      </PptxEditor.Root>
    ));
    expect(document.querySelector('.pptx-statusbar')!.getAttribute('aria-label')).toBe(
      'host:statusBar.ariaLabel'
    );
    expect(screen.getByRole('button', { name: 'host:statusBar.notes' })).toBeDefined();
    expect(screen.getByRole('button', { name: 'host:rail.collapse' })).toBeDefined();
    expect(screen.getByTestId('inside').textContent).toBe(
      'host:statusBar.notes|host:statusBar.notes|host'
    );
    expect(screen.getByTestId('outside').textContent).toBe('Notes|Notes|none');
  });

  test('a change of t keeps the compound mounted, with its state and focus', async () => {
    function Counter() {
      const [count, setCount] = useState(0);
      return (
        <button type="button" data-testid="count" onClick={() => setCount((n) => n + 1)}>
          {String(count)}
        </button>
      );
    }
    const ui = (t: ChromeTranslate | undefined) => (
      <PptxEditor.Root onReady={(e) => (ready = e)}>
        <PptxEditor.StatusBar t={t}>
          <Counter />
        </PptxEditor.StatusBar>
      </PptxEditor.Root>
    );
    let ready: PptxEditorInstance | null = null;
    const view = render(ui(host));
    await act(async () => void (await ready!.load(textDeck())));
    const button = screen.getByTestId('count');
    act(() => void fireEvent.click(button));
    button.focus();
    view.rerender(ui(undefined));
    expect(screen.getByTestId('count')).toBe(button);
    expect(button.textContent).toBe('1');
    expect(document.activeElement).toBe(button);
    expect(screen.getByRole('button', { name: 'Notes' })).toBeDefined();
    view.rerender(ui(host));
    expect(screen.getByTestId('count')).toBe(button);
    expect(screen.getByRole('button', { name: 'host:statusBar.notes' })).toBeDefined();
  });

  test('useChromeTranslate takes overrides', async () => {
    function Over() {
      const t = useChromeTranslate(new Map([['statusBar.notes', 'Speaker notes']]));
      return <output data-testid="over">{t('statusBar.notes')}</output>;
    }
    render(<Over />);
    expect(screen.getByTestId('over').textContent).toBe('Speaker notes');
  });

  test('StatusBar.Action is a host button after Notes', async () => {
    let sent = 0;
    await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <PptxEditor.StatusBar>
          <PptxEditor.StatusBar.Action label="Share" onSelect={() => (sent += 1)} />
        </PptxEditor.StatusBar>
      </PptxEditor.Root>
    ));
    const share = screen.getByRole('button', { name: 'Share' });
    const notes = screen.getByRole('button', { name: 'Notes' });
    expect(notes.compareDocumentPosition(share) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    act(() => void fireEvent.click(share));
    expect(sent).toBe(1);
  });
});

describe('useHistoryGroup', () => {
  test('the commands of one gesture are one undo step', async () => {
    let options: (() => { historyGroup?: unknown }) | null = null;
    function Slider() {
      const group = useHistoryGroup({ kind: 'range' });
      options = group.options;
      return <input type="range" data-testid="slider" ref={group.ref} />;
    }
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <Slider />
      </PptxEditor.Root>
    ));
    const slider = screen.getByTestId('slider');
    fireEvent.pointerDown(slider, { button: 0, pointerId: 1 });
    act(() => {
      editor.exec({ type: 'addSlide' }, options!() as never);
      editor.exec({ type: 'addSlide' }, options!() as never);
    });
    fireEvent.pointerUp(document, { pointerId: 1 });
    await new Promise((resolve) => setTimeout(resolve, 5));
    expect(editor.snapshot().slides.length).toBe(4);
    act(() => void editor.exec({ type: 'undo' }));
    expect(editor.snapshot().slides.length).toBe(2);
  });

  test('options() throws before the control mounts', () => {
    let group: ReturnType<typeof useHistoryGroup> | null = null;
    function Unbound() {
      group = useHistoryGroup({ kind: 'repeat' });
      return null;
    }
    render(<Unbound />);
    expect(() => group!.options()).toThrow();
  });
});

describe('useEditorCaret', () => {
  test('reads the caret of a text edit, and null otherwise', async () => {
    const seen = { caret: null as ReturnType<typeof useEditorCaret> };
    let renders = 0;
    function Probe() {
      seen.caret = useEditorCaret();
      renders += 1;
      return null;
    }
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <Probe />
      </PptxEditor.Root>
    ));
    expect(seen.caret).toBeNull();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const at = (offset: number) =>
      act(() =>
        editor.setSelection({
          kind: 'text',
          nodeId,
          anchor: { paragraph: 0, offset: 0 },
          focus: { paragraph: 0, offset },
        })
      );
    at(3);
    expect(seen.caret).toEqual({ nodeId, paragraph: 0, offset: 3 });
    const first = seen.caret;
    const before = renders;
    // The same focus with another anchor is the same caret: no new value, no render.
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 1 },
        focus: { paragraph: 0, offset: 3 },
      })
    );
    expect(seen.caret).toBe(first);
    expect(renders).toBe(before);
    act(() => editor.setSelection({ kind: 'none' }));
    expect(seen.caret).toBeNull();
  });
});

describe('useDocumentSearch', () => {
  test('finds, steps, goes to a match, and clears; hooks share the session', async () => {
    let search: ReturnType<typeof useDocumentSearch> | null = null;
    let other: ReturnType<typeof useDocumentSearch> | null = null;
    function Probe() {
      search = useDocumentSearch();
      return null;
    }
    function Other() {
      other = useDocumentSearch();
      return null;
    }
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <Probe />
        <Other />
      </PptxEditor.Root>
    ));
    act(() => search!.setQuery('one'));
    expect(search!.matches.length).toBe(3);
    expect(other!.query).toBe('one');
    expect(search!.activeIndex).toBe(-1);
    act(() => void search!.next());
    expect(search!.activeIndex).toBe(0);
    act(() => void search!.goTo(2));
    expect(search!.activeIndex).toBe(2);
    expect(search!.activeMatch?.slide).toBe(1);
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(search!.goTo(7)).toBe(false);
    act(() => search!.setWholeWord(true));
    act(() => search!.setQuery('on'));
    expect(search!.matches.length).toBe(0);
    act(() => search!.clear());
    expect(search!.query).toBe('');
    expect(search!.matches.length).toBe(0);
  });

  test('a hook that mounts again reads the matches of the deck as it is now', async () => {
    let search: ReturnType<typeof useDocumentSearch> | null = null;
    function Probe() {
      search = useDocumentSearch();
      return null;
    }
    let editor: PptxEditorInstance | null = null;
    const ui = (show: boolean) => (
      <PptxEditor.Root onReady={(e) => (editor = e)}>{show ? <Probe /> : null}</PptxEditor.Root>
    );
    const view = render(ui(true));
    await act(async () => void (await editor!.load(textDeck())));
    act(() => search!.setQuery('one'));
    expect(search!.matches.length).toBe(3);
    view.rerender(ui(false));
    // An edit with no hook mounted: the second slide goes, with its match.
    act(() => void editor!.exec({ type: 'deleteSlides', indices: [1] }));
    view.rerender(ui(true));
    expect(search!.query).toBe('one');
    expect(search!.matches.length).toBe(2);
  });

  test('outside a root it has no matches and does nothing', () => {
    let search: ReturnType<typeof useDocumentSearch> | null = null;
    function Probe() {
      search = useDocumentSearch();
      return null;
    }
    render(<Probe />);
    expect(search!.matches.length).toBe(0);
    expect(search!.next()).toBe(false);
    expect(search!.find('x').length).toBe(0);
  });
});
