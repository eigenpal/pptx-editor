/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The New comment row of the canvas context menus (`review.addComment`), after DOCX Editor's
// `packages/react/test/context-menu-add-comment.test.tsx`. The Vue twin is
// `../../vue/__tests__/context-menu-add-comment.test.ts`. DOCX Editor commit adf0e178.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { StrictMode, useState, type ReactNode } from 'react';
import type { EditorModule, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorViewport,
  SidePanel,
  useEditorCommand,
} from '@pptx-editor.dev/react';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxEditorReview, reviewModule } from '../index.ts';

const IN = 914_400;
const PX = 9_525;

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

async function mount(
  modules: readonly EditorModule[] | undefined,
  pane: ReactNode
): Promise<PptxEditorInstance> {
  let editor: PptxEditorInstance | null = null;
  render(
    <PptxEditorRoot onReady={(e) => (editor = e)} author="Ada Lovelace" modules={modules}>
      <PptxEditorViewport>
        <PptxEditorContent />
        {pane}
      </PptxEditorViewport>
    </PptxEditorRoot>
  );
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'Title', x: IN, y: IN, cx: 2 * IN, cy: IN }) }],
  });
  await act(async () => void (await editor!.load(deck)));
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  // happy-dom lays nothing out: one CSS pixel is one screen pixel at zoom 1.
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
  act(() => editor!.setZoom(1));
  return editor!;
}

/** Right-click the shape "Title", which selects it and opens the shape menu. */
function rightClickTitle(): void {
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  act(
    () => void fireEvent.contextMenu(stage, { clientX: (2 * IN) / PX, clientY: (1.5 * IN) / PX })
  );
  screen.getByRole('menu', { name: 'Shape actions' });
}

const row = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('[role="menu"] [data-slot="review.addComment"]');

const frame = (): Promise<void> => new Promise((done) => requestAnimationFrame(() => done()));

describe('New comment in the canvas context menu', () => {
  test('the row shows only while a comments pane is mounted with the review module', async () => {
    await mount([reviewModule()], <PptxEditorReview />);
    rightClickTitle();
    expect(row()).not.toBeNull();
    cleanup();
    await mount([reviewModule()], null);
    rightClickTitle();
    expect(row()).toBeNull();
    cleanup();
    // Without the review module the pane reads no comments: no row.
    await mount(undefined, <PptxEditorReview />);
    rightClickTitle();
    expect(row()).toBeNull();
  });

  test('the row opens a focused draft that points at the right-clicked element', async () => {
    const editor = await mount([reviewModule()], <PptxEditorReview />);
    rightClickTitle();
    act(() => void fireEvent.click(row()!));
    expect(screen.queryByRole('menu')).toBeNull();
    const field = screen.getByRole('textbox', { name: 'Add a comment…' });
    await act(frame);
    expect(document.activeElement).toBe(field);
    expect(screen.getByText('Points at Title')).toBeDefined();
    // The draft keeps its element when the selection moves on.
    act(() => editor.setSelection({ kind: 'none' }));
    act(() => void fireEvent.change(field, { target: { value: 'Check the title' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    const [thread] = editor.snapshot().comments!;
    expect(thread).toMatchObject({ text: 'Check the title', anchor: { kind: 'shape' } });
  });

  test('a closed comments pane opens on its tab with the draft', async () => {
    await mount(
      [reviewModule()],
      <SidePanel tab={undefined} panes={{ comments: <PptxEditorReview /> }} />
    );
    expect(document.querySelector('[data-tab="format"]')?.getAttribute('aria-selected')).toBe(
      'true'
    );
    rightClickTitle();
    act(() => void fireEvent.click(row()!));
    await act(frame);
    expect(document.querySelector('[data-tab="comments"]')?.getAttribute('aria-selected')).toBe(
      'true'
    );
    const field = screen.getByRole('textbox', { name: 'Add a comment…' });
    expect(document.activeElement).toBe(field);
  });

  test('a closed pane under StrictMode keeps the draft that its first mount opened', async () => {
    // StrictMode runs each effect twice: the second run must not close the draft.
    await mount(
      [reviewModule()],
      <StrictMode>
        <SidePanel tab={undefined} panes={{ comments: <PptxEditorReview /> }} />
      </StrictMode>
    );
    rightClickTitle();
    act(() => void fireEvent.click(row()!));
    await act(frame);
    expect(screen.getByRole('textbox', { name: 'Add a comment…' })).toBeDefined();
  });

  test('in view mode the row stays, off with the viewing reason', async () => {
    const editor = await mount([reviewModule()], <PptxEditorReview />);
    act(() => editor.setMode('view'));
    rightClickTitle();
    expect(row()?.getAttribute('aria-disabled')).toBe('true');
    act(() => void fireEvent.click(row()!));
    expect(screen.queryByRole('textbox', { name: 'Add a comment…' })).toBeNull();
  });
  test('a draft whose element is deleted posts on the slide', async () => {
    const editor = await mount([reviewModule()], <PptxEditorReview />);
    rightClickTitle();
    act(() => void fireEvent.click(row()!));
    expect(screen.getByText('Points at Title')).toBeDefined();
    act(() => void editor.exec({ type: 'delete' }));
    expect(screen.queryByText('Points at Title')).toBeNull();
    const field = screen.getByRole('textbox', { name: 'Add a comment…' });
    act(() => void fireEvent.change(field, { target: { value: 'On the slide' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    const [thread] = editor.snapshot().comments!;
    expect(thread).toMatchObject({ text: 'On the slide', anchor: { kind: 'slide' } });
  });

  test('a hidden pane opens no draft, so the row does not show', async () => {
    await mount([reviewModule()], <PptxEditorReview hidden />);
    rightClickTitle();
    expect(row()).toBeNull();
  });

  test('useEditorCommand follows the pane as it mounts and unmounts', async () => {
    let show: (on: boolean) => void = () => {};
    function Probe() {
      const command = useEditorCommand('review.addComment');
      return <output data-testid="probe">{String(command.isEnabled)}</output>;
    }
    function Toggle() {
      const [on, setOn] = useState(false);
      show = setOn;
      return (
        <>
          <Probe />
          {on ? <PptxEditorReview /> : null}
        </>
      );
    }
    await mount([reviewModule()], <Toggle />);
    const probe = (): string | null => screen.getByTestId('probe').textContent;
    expect(probe()).toBe('false');
    act(() => show(true));
    expect(probe()).toBe('true');
    act(() => show(false));
    expect(probe()).toBe('false');
  });
});
