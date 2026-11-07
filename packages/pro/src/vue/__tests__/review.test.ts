/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The twin of `../../react/__tests__/review.test.tsx`.
import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, defineComponent, h, nextTick, type App, type Component, type VNode } from 'vue';
import {
  listenSlideshowStartKeys,
  type EditorModule,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  PptxEditorContent,
  PptxEditorRoot,
  PptxEditorSidePanel,
  PptxEditorViewport,
} from '@pptx-editor.dev/vue';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxEditorReview, reviewModule, useReview, type UseReviewReturn } from '../index.ts';
import { repeatedThreadDeck } from '../../__tests__/repeat-deck.ts';

const IN = 914_400;

let app: App | null = null;
afterEach(() => {
  app?.unmount();
  app = null;
  document.body.replaceChildren();
});

async function mount(
  modules: readonly EditorModule[] | undefined,
  pane: () => VNode = () => h(PptxEditorReview)
): Promise<PptxEditorInstance> {
  const host = document.createElement('div');
  document.body.append(host);
  let ready: PptxEditorInstance | null = null;
  app = createApp(
    defineComponent({
      setup: () => () =>
        h(
          PptxEditorRoot,
          { onReady: (e: PptxEditorInstance) => (ready = e), author: 'Ada Lovelace', modules },
          () => h(PptxEditorViewport, null, () => [h(PptxEditorContent), pane()])
        ),
    }) as Component
  );
  app.mount(host);
  await nextTick();
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'Title', x: IN, y: IN, cx: IN, cy: IN }) }],
  });
  const editor = ready as PptxEditorInstance | null;
  expect((await editor!.load(deck)).ok).toBe(true);
  await settle();
  return editor!;
}

const settle = async (): Promise<void> => {
  for (let i = 0; i < 3; i += 1) await nextTick();
};
const button = (name: string): HTMLButtonElement | undefined =>
  [...document.querySelectorAll<HTMLButtonElement>('button')].find(
    (b) => (b.getAttribute('aria-label') ?? b.textContent?.trim()) === name
  );
const field = (label: string): HTMLTextAreaElement | null =>
  document.querySelector<HTMLTextAreaElement>(`textarea[aria-label="${label}"]`);

async function type(el: HTMLTextAreaElement, text: string): Promise<void> {
  el.value = text;
  el.dispatchEvent(new Event('input', { bubbles: true }));
  await settle();
}
async function key(el: HTMLElement, init: KeyboardEventInit): Promise<void> {
  el.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
  await settle();
}
async function click(el: HTMLElement | undefined): Promise<void> {
  el!.click();
  await settle();
}
async function post(text: string): Promise<void> {
  await click(button('New comment'));
  const el = field('Add a comment…')!;
  await type(el, text);
  await key(el, { key: 'Enter' });
}

describe('PptxEditorReview (Vue)', () => {
  test('without a review module it says that comments are kept', async () => {
    await mount(undefined);
    expect(document.body.textContent).toContain('Comments are kept when you save.');
    expect(button('New comment')).toBeUndefined();
  });

  test('adds, replies to, resolves, reopens, and deletes a thread', async () => {
    const editor = await mount([reviewModule()]);
    await post('Check <b>this</b>');
    const [thread] = editor.snapshot().comments!;
    expect(thread).toMatchObject({ author: 'Ada Lovelace', text: 'Check <b>this</b>' });
    const list = document.querySelector('ul[aria-label="Comments on this slide"]')!;
    expect(list.querySelectorAll(':scope > li')).toHaveLength(1);
    // File text renders as text.
    expect(list.querySelector('b')).toBeNull();
    expect(list.textContent).toContain('Check <b>this</b>');

    editor.selectComment(thread!.id);
    await settle();
    expect(list.querySelector('[aria-current="true"]')).not.toBeNull();
    const reply = field('Reply to Ada Lovelace')!;
    await type(reply, 'Done');
    await key(reply, { key: 'Enter' });
    expect(editor.snapshot().comments![0]!.replies.map((r) => r.text)).toEqual(['Done']);

    await click(button('Resolve thread by Ada Lovelace'));
    expect(editor.snapshot().comments![0]!.resolved).toBe(true);
    expect(document.body.textContent).toContain('Resolved');
    await click(button('Reopen thread by Ada Lovelace'));
    expect(editor.snapshot().comments![0]!.resolved).toBe(false);

    await click(button('Delete comment by Ada Lovelace'));
    expect(editor.snapshot().comments).toEqual([]);
    expect(document.body.textContent).toContain('No comments on this slide');
  });

  test('Shift+Enter keeps the field open; Escape cancels the draft', async () => {
    const editor = await mount([reviewModule()]);
    await click(button('New comment'));
    const el = field('Add a comment…')!;
    await type(el, 'Line');
    await key(el, { key: 'Enter', shiftKey: true });
    expect(editor.snapshot().comments).toEqual([]);
    await key(el, { key: 'Escape' });
    expect(field('Add a comment…')).toBeNull();
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
      await click(button('New comment'));
      const el = field('Add a comment…')!;
      await type(el, 'Note');
      await key(el, { key: 'Enter', ctrlKey: true });
      expect(started).toBe(0);
      expect(editor.snapshot().comments!.map((c) => c.text)).toEqual(['Note']);
      // The same key outside a comment field starts the show.
      await key(document.body, { key: 'Enter', ctrlKey: true });
      expect(started).toBe(1);
    } finally {
      stop();
    }
  });

  test('a field whose comment is gone closes and saves nothing', async () => {
    const editor = await mount([reviewModule()]);
    expect((await editor.load(await repeatedThreadDeck())).ok).toBe(true);
    await settle();
    const [a] = editor.snapshot().comments!;
    editor.selectComment(a!.id);
    await settle();
    await type(field('Reply to Ada Lovelace')!, 'For A');
    await click(button('Edit comment by Ada Lovelace'));
    await type(field('Edit comment')!, 'A edited');
    // A peer deletes A: B, a repeat with A's author and date, takes A's id and card.
    editor.exec({ type: 'deleteComment', commentId: a!.id });
    await settle();
    expect(editor.snapshot().comments!.map((t) => [t.id === a!.id, t.text])).toEqual([
      [true, 'B'],
      [false, 'C'],
    ]);
    expect(field('Edit comment')).toBeNull();
    expect(field('Reply to Ada Lovelace')!.value).toBe('');
    expect(editor.snapshot().comments![0]!.replies).toEqual([]);
  });

  test('parts: hidden removes a part, a card class reaches each card, filter picks threads', async () => {
    const editor = await mount([reviewModule()], () =>
      h(
        PptxEditorReview,
        { card: { className: 'host-card' }, filter: (i: { resolved: boolean }) => !i.resolved },
        () => [h(PptxEditorReview.Delete, { hidden: true })]
      )
    );
    await post('One');
    await post('Two');
    expect(document.querySelectorAll('.host-card')).toHaveLength(2);
    expect(button('Delete comment by Ada Lovelace')).toBeUndefined();
    const [first] = editor.snapshot().comments!;
    editor.exec({ type: 'resolveComment', commentId: first!.id });
    await settle();
    expect(document.querySelectorAll('.host-card')).toHaveLength(1);
  });

  test('preset=false renders only the parts you pass; asChild wires your button', async () => {
    await mount([reviewModule()], () =>
      h(PptxEditorReview, { preset: false }, () => [
        h(PptxEditorReview.AddComment, { asChild: true }, () => [
          h('button', { type: 'button' }, 'Ask'),
        ]),
        h(PptxEditorReview.Draft),
      ])
    );
    expect(button('New comment')).toBeUndefined();
    expect(document.querySelector('ul')).toBeNull();
    await click(button('Ask'));
    expect(field('Add a comment…')).not.toBeNull();
  });

  test('useReview gives the data and the actions', async () => {
    let review: UseReviewReturn | null = null;
    const Probe = defineComponent({
      setup() {
        review = useReview();
        return () => null;
      },
    });
    await mount([reviewModule()], () => h(Probe));
    expect(review!.ready.value).toBe(true);
    expect(review!.author.value).toBe('Ada Lovelace');
    review!.comment('From a composable');
    await settle();
    expect(review!.items.value.map((i) => i.text)).toEqual(['From a composable']);
  });

  test('the side panel shows the pane in its #comments slot', async () => {
    await mount([reviewModule()], () =>
      h(PptxEditorSidePanel, { tab: 'comments' }, { comments: () => h(PptxEditorReview) })
    );
    expect(button('New comment')).toBeDefined();
  });
});
