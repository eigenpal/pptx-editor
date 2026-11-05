// The hooks a host builds its own chrome with: `useZoom`, `useEditorEvent`,
// `useEditorSnapshot`, and `useScopeClassName`. The Vue twin is
// packages/vue/test/customization-hooks.test.ts.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { useState, type ReactNode } from 'react';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { resetDevWarnings, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import {
  PptxEditor,
  useEditorEvent,
  useEditorSnapshot,
  useScopeClassName,
  useZoom,
  usePptxEditor,
} from '../index.ts';

afterEach(() => cleanup());

const deck = () => buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] });

async function mount(ui: (ready: (e: PptxEditorInstance) => void) => ReactNode) {
  let editor: PptxEditorInstance | null = null;
  render(<>{ui((e) => (editor = e))}</>);
  await act(async () => void (await editor!.load(deck())));
  return editor!;
}

function ZoomProbe() {
  const z = useZoom();
  return (
    <div>
      <output data-testid="zoom">{`${Math.round(z.zoom * 100)} ${z.isFit ? 'fit' : 'fixed'}`}</output>
      <button onClick={z.zoomIn} disabled={!z.canZoomIn}>
        in
      </button>
      <button onClick={z.zoomOut} disabled={!z.canZoomOut}>
        out
      </button>
      <button onClick={z.fit}>fit</button>
      <button onClick={z.reset}>reset</button>
      <button onClick={() => z.setZoom(9)}>huge</button>
    </div>
  );
}

const shown = (): string => screen.getByTestId('zoom').textContent ?? '';

describe('useZoom', () => {
  test('reads the zoom and sets it, steps it, fits, and resets', async () => {
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <ZoomProbe />
      </PptxEditor.Root>
    ));
    expect(shown().endsWith('fit')).toBe(true);
    act(() => void fireEvent.click(screen.getByText('reset')));
    expect(shown()).toBe('100 fixed');
    act(() => void fireEvent.click(screen.getByText('in')));
    expect(shown()).toBe('110 fixed');
    act(() => void fireEvent.click(screen.getByText('out')));
    act(() => void fireEvent.click(screen.getByText('out')));
    expect(editor.snapshot().zoom).toBe(0.9);
    act(() => void fireEvent.click(screen.getByText('huge')));
    expect(shown()).toBe('400 fixed');
    expect((screen.getByText('in') as HTMLButtonElement).disabled).toBe(true);
    act(() => void fireEvent.click(screen.getByText('fit')));
    expect(editor.snapshot().fitZoom).toBe(true);
    act(() => editor.setZoom(0.5));
    expect(shown()).toBe('50 fixed');
  });

  test('outside a root it reports 100 % and its actions do nothing', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    render(<ZoomProbe />);
    expect(shown()).toBe('100 fit');
    expect((screen.getByText('in') as HTMLButtonElement).disabled).toBe(true);
    act(() => void fireEvent.click(screen.getByText('reset')));
    expect(shown()).toBe('100 fit');
    warn.mockRestore();
  });
});

describe('useEditorEvent', () => {
  test('calls the latest handler for its event and stops on unmount', async () => {
    const log: string[] = [];
    function Listener({ tag }: { tag: string }) {
      useEditorEvent('selectionChange', (s) => void log.push(`${tag}:${s.currentSlide}`));
      useEditorEvent('rejection', (r) => void log.push(`rejection:${r.reason}`));
      return null;
    }
    let setTag: (tag: string) => void = () => undefined;
    let setShown: (on: boolean) => void = () => undefined;
    function Host() {
      const [tag, tagSetter] = useState('a');
      const [on, onSetter] = useState(true);
      setTag = tagSetter;
      setShown = onSetter;
      return on ? <Listener tag={tag} /> : null;
    }
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <Host />
      </PptxEditor.Root>
    ));
    log.length = 0;
    act(() => editor.selectSlide(1));
    act(() => setTag('b'));
    act(() => editor.selectSlide(0));
    act(() => void editor.exec({ type: 'undo' }));
    expect(log).toEqual(['a:1', 'b:0', 'rejection:disabled.nothingToUndo']);
    act(() => setShown(false));
    act(() => editor.selectSlide(1));
    expect(log.length).toBe(3);
  });
});

describe('useEditorSnapshot', () => {
  test('changes its counter on each editor change', async () => {
    const seen: number[] = [];
    function Probe() {
      seen.push(useEditorSnapshot(usePptxEditor()));
      return null;
    }
    const editor = await mount((onReady) => (
      <PptxEditor.Root onReady={onReady}>
        <Probe />
      </PptxEditor.Root>
    ));
    const before = seen.at(-1)!;
    act(() => editor.selectSlide(1));
    expect(seen.at(-1)! > before).toBe(true);
  });
});

describe('useScopeClassName', () => {
  function Part() {
    return <aside data-testid="part" className={`${useScopeClassName()}my-panel`} />;
  }
  const part = (): string => screen.getByTestId('part').className;

  test('outside the viewport it is the scope class, with dark for a dark root', () => {
    render(
      <PptxEditor.Root>
        <Part />
      </PptxEditor.Root>
    );
    expect(part()).toBe('pptx-editor my-panel');
    cleanup();
    render(
      <PptxEditor.Root colorMode="dark">
        <Part />
      </PptxEditor.Root>
    );
    expect(part()).toBe('pptx-editor dark my-panel');
  });

  test('inside the viewport it is empty: the viewport carries the scope', () => {
    render(
      <PptxEditor.Root colorMode="dark">
        <PptxEditor.Viewport>
          <Part />
        </PptxEditor.Viewport>
      </PptxEditor.Root>
    );
    expect(part()).toBe('my-panel');
    expect(screen.getByTestId('part').closest('.pptx-editor.dark') !== null).toBe(true);
  });
});
