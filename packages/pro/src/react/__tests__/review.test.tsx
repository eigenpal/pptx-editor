/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/** @jsxImportSource react */
import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactNode } from 'react';
import {
  listenSlideshowStartKeys,
  type EditorModule,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorViewport,
  SidePanel,
} from '@pptx-editor.dev/react';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxEditorReview, reviewModule, useReview, useReviewAuthor } from '../index.ts';

const IN = 914_400;

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

async function mount(
  modules: readonly EditorModule[] | undefined,
  pane: ReactNode = <PptxEditorReview />
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
    slides: [{ shapes: shape({ id: 2, name: 'Title', x: IN, y: IN, cx: IN, cy: IN }) }],
  });
  await act(async () => void (await editor!.load(deck)));
  return editor!;
}

function post(text: string): void {
  act(() => void fireEvent.click(screen.getByRole('button', { name: 'New comment' })));
  const field = screen.getByRole('textbox', { name: 'Add a comment…' });
  expect(document.activeElement).toBe(field);
  act(() => void fireEvent.change(field, { target: { value: text } }));
  act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
}

describe('PptxEditorReview', () => {
  test('without a review module it says that comments are kept', async () => {
    await mount(undefined);
    expect(screen.getByText('Comments are kept when you save.')).toBeDefined();
    expect(screen.queryByRole('button', { name: 'New comment' })).toBeNull();
  });

  test('adds, replies to, resolves, reopens, and deletes a thread', async () => {
    const editor = await mount([reviewModule()]);
    post('Check <b>this</b>');
    const [thread] = editor.snapshot().comments!;
    expect(thread).toMatchObject({ author: 'Ada Lovelace', text: 'Check <b>this</b>' });
    // File text renders as text.
    expect(screen.getByText('Check <b>this</b>').querySelector('b')).toBeNull();
    const list = screen.getByRole('list', { name: 'Comments on this slide' });
    expect(list.querySelectorAll(':scope > li')).toHaveLength(1);

    // A pick from elsewhere (a marker) marks the card and shows its reply line.
    act(() => editor.selectComment(thread!.id));
    expect(list.querySelector('[aria-current="true"]')).not.toBeNull();
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: 'Done' } }));
    act(() => void fireEvent.keyDown(reply, { key: 'Enter' }));
    expect(editor.snapshot().comments![0]!.replies.map((r) => r.text)).toEqual(['Done']);

    const resolve = screen.getByRole('button', { name: 'Resolve thread by Ada Lovelace' });
    act(() => void fireEvent.click(resolve));
    expect(editor.snapshot().comments![0]!.resolved).toBe(true);
    expect(screen.getByText('Resolved')).toBeDefined();
    act(
      () =>
        void fireEvent.click(screen.getByRole('button', { name: 'Reopen thread by Ada Lovelace' }))
    );
    expect(editor.snapshot().comments![0]!.resolved).toBe(false);

    act(() => {
      const buttons = screen.getAllByRole('button', { name: 'Delete comment by Ada Lovelace' });
      fireEvent.click(buttons[0]!);
    });
    expect(editor.snapshot().comments).toEqual([]);
    expect(screen.getByText('No comments on this slide')).toBeDefined();
  });

  test('Shift+Enter keeps the field open; Escape cancels the draft', async () => {
    const editor = await mount([reviewModule()]);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'New comment' })));
    const field = screen.getByRole('textbox', { name: 'Add a comment…' });
    act(() => void fireEvent.change(field, { target: { value: 'Line' } }));
    act(() => void fireEvent.keyDown(field, { key: 'Enter', shiftKey: true }));
    expect(editor.snapshot().comments).toEqual([]);
    act(() => void fireEvent.keyDown(field, { key: 'Escape' }));
    expect(screen.queryByRole('textbox', { name: 'Add a comment…' })).toBeNull();
    expect(editor.snapshot().comments).toEqual([]);
  });

  test('Enter in a comment field posts and does not start the slide show', async () => {
    const editor = await mount([reviewModule()]);
    let started = 0;
    const stop = listenSlideshowStartKeys(
      () => document.querySelector('.pptx-review'),
      () => false,
      () => void started++
    );
    try {
      post('Note');
      expect(started).toBe(0);
      expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['Note']);
    } finally {
      stop();
    }
  });

  test('parts: hidden removes a part, a card class reaches each card, filter picks threads', async () => {
    const editor = await mount(
      [reviewModule()],
      <PptxEditorReview card={{ className: 'host-card' }} filter={(i) => !i.resolved}>
        <PptxEditorReview.Delete hidden />
      </PptxEditorReview>
    );
    post('One');
    post('Two');
    expect(document.querySelectorAll('.host-card')).toHaveLength(2);
    expect(screen.queryByRole('button', { name: 'Delete comment by Ada Lovelace' })).toBeNull();
    const [first] = editor.snapshot().comments!;
    act(() => void editor.exec({ type: 'resolveComment', commentId: first!.id }));
    expect(document.querySelectorAll('.host-card')).toHaveLength(1);
  });

  test('preset={false} renders only the parts you pass; asChild wires your button', async () => {
    await mount(
      [reviewModule()],
      <PptxEditorReview preset={false}>
        <PptxEditorReview.AddComment asChild>
          <button type="button">Ask</button>
        </PptxEditorReview.AddComment>
        <PptxEditorReview.Draft />
      </PptxEditorReview>
    );
    expect(screen.queryByRole('button', { name: 'New comment' })).toBeNull();
    expect(screen.queryByRole('list')).toBeNull();
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Ask' })));
    expect(screen.getByRole('textbox', { name: 'Add a comment…' })).toBeDefined();
  });

  test('useReview and useReviewAuthor give the data and the actions', async () => {
    let review: ReturnType<typeof useReview> | null = null;
    let me: ReturnType<typeof useReviewAuthor> | null = null;
    function Probe() {
      review = useReview();
      me = useReviewAuthor('Ada Lovelace');
      return null;
    }
    await mount([reviewModule()], <Probe />);
    expect(review!.ready).toBe(true);
    expect(review!.author).toBe('Ada Lovelace');
    expect(me!.color).toBe(`var(--doc-review-author-${me!.slot})`);
    act(() => void review!.comment('From a hook'));
    expect(review!.items.map((i) => i.text)).toEqual(['From a hook']);
  });

  test('SidePanel shows the pane in its Comments tab', async () => {
    await mount(
      [reviewModule()],
      <SidePanel tab="comments" panes={{ comments: <PptxEditorReview /> }} />
    );
    expect(screen.getByRole('button', { name: 'New comment' })).toBeDefined();
  });
});
