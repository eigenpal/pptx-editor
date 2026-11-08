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
  EditNotice,
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorViewport,
  SidePanel,
} from '@pptx-editor.dev/react';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxEditorReview, reviewModule, useReview, useReviewAuthor } from '../index.ts';
import { repeatedThreadDeck } from '../../__tests__/repeat-deck.ts';

const IN = 914_400;

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

async function mount(
  modules: readonly EditorModule[] | undefined,
  pane: ReactNode = <PptxEditorReview />,
  text?: string
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
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Title',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: IN,
          ...(text ? { text: para(run(text)) } : {}),
        }),
      },
    ],
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

  test('Ctrl+Enter in a comment field posts and does not start the slide show', async () => {
    const editor = await mount([reviewModule()]);
    let started = 0;
    const stop = listenSlideshowStartKeys(
      () => document.querySelector('.pptx-editor'),
      () => false,
      () => void started++
    );
    try {
      act(() => void fireEvent.click(screen.getByRole('button', { name: 'New comment' })));
      const field = screen.getByRole('textbox', { name: 'Add a comment…' });
      act(() => void fireEvent.change(field, { target: { value: 'Note' } }));
      act(() => void fireEvent.keyDown(field, { key: 'Enter', ctrlKey: true }));
      expect(started).toBe(0);
      expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['Note']);
      // The same key outside a comment field starts the show.
      act(() => void fireEvent.keyDown(document.body, { key: 'Enter', ctrlKey: true }));
      expect(started).toBe(1);
    } finally {
      stop();
    }
  });

  test('a field whose comment is gone closes and saves nothing', async () => {
    const editor = await mount([reviewModule()]);
    await act(async () => void (await editor.load(await repeatedThreadDeck())));
    const [a] = editor.snapshot().comments!;
    act(() => editor.selectComment(a!.id));
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: 'For A' } }));
    const edit = screen.getAllByRole('button', { name: 'Edit comment by Ada Lovelace' })[0]!;
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' });
    act(() => void fireEvent.change(box, { target: { value: 'A edited' } }));
    // A peer deletes A: B, a repeat with A's author and date, keeps its own card.
    act(() => void editor.exec({ type: 'deleteComment', commentId: a!.id }));
    expect(editor.snapshot().comments!.map((t) => [t.id === a!.id, t.text])).toEqual([
      [false, 'B'],
      [false, 'C'],
    ]);
    expect(screen.queryByRole('textbox', { name: 'Edit comment' })).toBeNull();
    // The pick ends with A, so no reply line carries A's draft to B.
    expect(editor.snapshot().activeComment).toBeNull();
    expect(screen.queryByRole('textbox', { name: 'Reply to Ada Lovelace' })).toBeNull();
    expect(editor.snapshot().comments![0]!.replies).toEqual([]);
  });

  test("a peer's edit of the thread keeps the unsent reply and the open edit box", async () => {
    const editor = await mount([reviewModule()]);
    post('First');
    const [thread] = editor.snapshot().comments!;
    act(() => editor.selectComment(thread!.id));
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: 'Draft' } }));
    const edit = screen.getByRole('button', { name: 'Edit comment by Ada Lovelace' });
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' });
    act(() => void fireEvent.change(box, { target: { value: 'Mine' } }));
    act(() => void editor.exec({ type: 'editComment', commentId: thread!.id, text: 'Theirs' }));
    expect(editor.snapshot().comments![0]!.text).toBe('Theirs');
    const value = (name: string): string =>
      (screen.getByRole('textbox', { name }) as HTMLTextAreaElement).value;
    expect(value('Edit comment')).toBe('Mine');
    expect(value('Reply to Ada Lovelace')).toBe('Draft');
    act(
      () =>
        void fireEvent.keyDown(screen.getByRole('textbox', { name: 'Edit comment' }), {
          key: 'Enter',
        })
    );
    expect(editor.snapshot().comments![0]!.text).toBe('Mine');
  });

  test('a delete of a twin keeps the fields and the pick of another twin', async () => {
    const editor = await mount([reviewModule()]);
    await act(async () => void (await editor.load(await repeatedThreadDeck())));
    const [a, b] = editor.snapshot().comments!;
    act(() => editor.selectComment(b!.id));
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: 'For B' } }));
    const edit = screen.getAllByRole('button', { name: 'Edit comment by Ada Lovelace' })[1]!;
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' });
    act(() => void fireEvent.change(box, { target: { value: 'B edited' } }));
    act(() => void editor.exec({ type: 'deleteComment', commentId: a!.id }));
    expect(editor.snapshot().comments![0]).toMatchObject({ text: 'B', identity: b!.identity });
    expect(editor.snapshot().activeComment).toBe(editor.snapshot().comments![0]!.id);
    const value = (name: string): string =>
      (screen.getByRole('textbox', { name }) as HTMLTextAreaElement).value;
    expect(value('Edit comment')).toBe('B edited');
    expect(value('Reply to Ada Lovelace')).toBe('For B');
  });

  test('a save on a twin that a peer changed is refused: the box keeps the text and says so', async () => {
    const editor = await mount([reviewModule()]);
    await act(async () => void (await editor.load(await repeatedThreadDeck())));
    const [a] = editor.snapshot().comments!;
    const edit = screen.getAllByRole('button', { name: 'Edit comment by Ada Lovelace' })[0]!;
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' });
    act(() => void fireEvent.change(box, { target: { value: 'Mine' } }));
    act(() => void editor.exec({ type: 'editComment', commentId: a!.id, text: 'Peer' }));
    const field = screen.getByRole('textbox', { name: 'Edit comment' }) as HTMLTextAreaElement;
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['Peer', 'B', 'C']);
    expect(field.value).toBe('Mine');
    const notice = document.querySelector('.pptx-review__notice')!;
    expect(notice.getAttribute('role')).toBe('status');
    expect(notice.firstElementChild!.textContent).toBe(
      'This comment changed. Save again to replace it.'
    );
    expect(notice.querySelector('.pptx-review__current')!.textContent).toBe('It now reads: Peer');
    expect(field.getAttribute('aria-describedby')).toContain(notice.id);
    // A second, deliberate save replaces the comment as it shows now.
    act(() => void fireEvent.keyDown(field, { key: 'Enter' }));
    expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['Mine', 'B', 'C']);
    expect(screen.queryByRole('textbox', { name: 'Edit comment' })).toBeNull();
  });

  test('a refusal for another reason says that reason, and the notice clears on typing', async () => {
    const editor = await mount([reviewModule()]);
    post('First');
    act(
      () =>
        void fireEvent.click(screen.getByRole('button', { name: 'Edit comment by Ada Lovelace' }))
    );
    const box = screen.getByRole('textbox', { name: 'Edit comment' }) as HTMLTextAreaElement;
    act(() => void fireEvent.change(box, { target: { value: '\u0001' } }));
    act(() => void fireEvent.keyDown(box, { key: 'Enter' }));
    const notice = document.querySelector('.pptx-review__notice')!;
    expect(notice.textContent).toBe('Comment text must have 1 to 10,000 characters');
    expect(editor.snapshot().comments![0]!.text).toBe('First');
    act(() => void fireEvent.change(box, { target: { value: 'Second' } }));
    expect(notice.textContent).toBe('');
    expect(notice.isConnected).toBe(true);
  });

  test('without an author, Enter in the reply line posts nothing and says nothing new', async () => {
    const editor = await mount([reviewModule()]);
    post('First');
    act(() => editor.setAuthor(undefined));
    const [thread] = editor.snapshot().comments!;
    act(() => editor.selectComment(thread!.id));
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: 'Hello' } }));
    act(() => void fireEvent.keyDown(reply, { key: 'Enter' }));
    expect(editor.snapshot().comments![0]!.replies).toEqual([]);
    expect((reply as HTMLTextAreaElement).value).toBe('Hello');
    expect(document.querySelector('.pptx-review__notice')!.textContent).toBe('');
    expect(screen.getByText('Set an author to comment.')).toBeDefined();
  });

  test('the pane edits the copy of an exact duplicate that it shows, and says a refusal once', async () => {
    const editor = await mount(
      [reviewModule()],
      <>
        <PptxEditorReview />
        <EditNotice />
      </>
    );
    await act(async () => void (await editor.load(await repeatedThreadDeck('same'))));
    const edit = screen.getAllByRole('button', { name: 'Edit comment by Ada Lovelace' })[1]!;
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' }) as HTMLTextAreaElement;
    act(() => void fireEvent.change(box, { target: { value: '\u0001' } }));
    act(() => void fireEvent.keyDown(box, { key: 'Enter' }));
    expect(document.querySelector('.pptx-review__notice')!.textContent).toBe(
      'Comment text must have 1 to 10,000 characters'
    );
    // The field shows the refusal: the edit notice does not repeat it.
    expect(document.querySelector('.pptx-edit-notice')).toBeNull();
    // The identity names the copy, which text alone cannot.
    act(() => void fireEvent.change(box, { target: { value: 'Mine' } }));
    act(() => void fireEvent.keyDown(box, { key: 'Enter' }));
    expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['A', 'Mine', 'A']);
  });

  test('the draft shows its refusal, and a reply notice goes with the pick', async () => {
    const editor = await mount([reviewModule()]);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'New comment' })));
    const draft = screen.getByRole('textbox', { name: 'Add a comment…' });
    act(() => void fireEvent.change(draft, { target: { value: '\u0001' } }));
    act(() => void fireEvent.keyDown(draft, { key: 'Enter' }));
    expect(document.querySelector('.pptx-review__notice')!.textContent).toBe(
      'Comment text must have 1 to 10,000 characters'
    );
    act(() => void fireEvent.keyDown(draft, { key: 'Escape' }));
    post('First');
    const [thread] = editor.snapshot().comments!;
    act(() => editor.selectComment(thread!.id));
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: '\u0001' } }));
    act(() => void fireEvent.keyDown(reply, { key: 'Enter' }));
    const notice = () => document.querySelector('.pptx-review__reply-line .pptx-review__notice');
    expect(notice()!.textContent).not.toBe('');
    act(() => editor.selectComment(null));
    act(() => editor.selectComment(thread!.id));
    expect(notice()!.textContent).toBe('');
  });

  test('when one of three copies goes, an edit of a copy closes and saves nothing', async () => {
    const editor = await mount([reviewModule()]);
    await act(async () => void (await editor.load(await repeatedThreadDeck('same'))));
    const [first] = editor.snapshot().comments!;
    const edit = screen.getAllByRole('button', { name: 'Edit comment by Ada Lovelace' })[0]!;
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' });
    act(() => void fireEvent.change(box, { target: { value: 'Mine' } }));
    act(() => void editor.exec({ type: 'deleteComment', commentId: first!.id }));
    expect(screen.queryByRole('textbox', { name: 'Edit comment' })).toBeNull();
    expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['A', 'A']);
  });

  test('a delete that changes a thread id keeps its card, its pick, and its fields', async () => {
    const editor = await mount([reviewModule()]);
    await act(async () => void (await editor.load(await repeatedThreadDeck('pair'))));
    const [a, b] = editor.snapshot().comments!;
    act(() => editor.selectComment(b!.id));
    const reply = screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' });
    act(() => void fireEvent.change(reply, { target: { value: 'For B' } }));
    const edit = screen.getAllByRole('button', { name: 'Edit comment by Ada Lovelace' })[1]!;
    act(() => void fireEvent.click(edit));
    const box = screen.getByRole('textbox', { name: 'Edit comment' });
    act(() => void fireEvent.change(box, { target: { value: 'B edited' } }));
    // A peer deletes A: B is the last copy of the GUID, so its id loses the mark.
    act(() => void editor.exec({ type: 'deleteComment', commentId: a!.id }));
    const [left] = editor.snapshot().comments!;
    expect(left).toMatchObject({ text: 'B', identity: b!.identity });
    expect(left!.id).not.toBe(b!.id);
    expect(editor.snapshot().activeComment).toBe(left!.id);
    const value = (name: string): string =>
      (screen.getByRole('textbox', { name }) as HTMLTextAreaElement).value;
    expect(value('Edit comment')).toBe('B edited');
    expect(value('Reply to Ada Lovelace')).toBe('For B');
    act(
      () =>
        void fireEvent.keyDown(screen.getByRole('textbox', { name: 'Reply to Ada Lovelace' }), {
          key: 'Enter',
        })
    );
    expect(editor.snapshot().comments![0]!.replies.map((r) => r.text)).toEqual(['For B']);
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

  test("cards draw each author in the editor's roster color, not a name hash", async () => {
    const editor = await mount([reviewModule()]);
    act(() => void editor.exec({ type: 'addComment', text: 'One', author: 'Grace Hopper' }));
    act(() => void editor.exec({ type: 'addComment', text: 'Two' }));
    const roster = editor.getReviewAuthors();
    // The signing author holds the first slot from the start; Grace takes the next one.
    expect(roster.map((a) => [a.author, a.slot])).toEqual([
      ['Ada Lovelace', 0],
      ['Grace Hopper', 1],
    ]);
    const list = screen.getByRole('list', { name: 'Comments on this slide' });
    for (const info of roster) {
      const card = list.querySelector(`[data-review-author="${info.author}"]`) as HTMLElement;
      expect(card.getAttribute('data-review-author-slot')).toBe(String(info.slot));
      expect(card.style.getPropertyValue('--doc-review-author-current')).toBe(info.color);
    }
    let grace: ReturnType<typeof useReviewAuthor> | null = null;
    function Probe() {
      grace = useReviewAuthor('Grace Hopper');
      return null;
    }
    cleanup();
    await mount([reviewModule()], <Probe />);
    expect(grace).toBeUndefined();
  });

  test('a draft on selected text points the comment at that text', async () => {
    const editor = await mount([reviewModule()], <PptxEditorReview />, 'Quick brown fox');
    const nodeId = editor.snapshot().shapes![0]!.nodeId;
    act(() =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 6 },
        focus: { paragraph: 0, offset: 11 },
      })
    );
    post('Which brown?');
    expect(editor.snapshot().comments![0]!.anchor).toEqual({
      kind: 'text',
      nodeId,
      start: { paragraph: 0, offset: 6 },
      end: { paragraph: 0, offset: 11 },
    });
  });

  test('SidePanel shows the pane in its Comments tab', async () => {
    await mount(
      [reviewModule()],
      <SidePanel tab="comments" panes={{ comments: <PptxEditorReview /> }} />
    );
    expect(screen.getByRole('button', { name: 'New comment' })).toBeDefined();
  });
});
