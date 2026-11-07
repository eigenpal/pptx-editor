/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The comments of the current slide, as a composable: the twin of `../react/useReview.ts`, after
// DOCX Editor's Vue `useReview` (docx-editor/packages/pro/src/vue/useReview.ts at adf0e178c).

import { computed, shallowRef, toValue, watch, type ComputedRef, type MaybeRefOrGetter } from 'vue';
import type { EditorSnapshot, PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { usePptxEditor } from '@pptx-editor.dev/vue';
import {
  commentTarget,
  reviewActions,
  reviewItems,
  reviewReadOnly,
  reviewReady,
  type ReviewActivationOptions,
  type ReviewCommentTarget,
  type ReviewItemView,
} from '../review/comment-view.ts';

/**
 * What `useReview` returns: the twin of React's `UseReviewReturn`, with the state as computed
 * refs. Every action reports whether the editor applied it.
 *
 * @public
 */
export interface UseReviewReturn {
  /** The comments of the current slide, in file order: each thread, then its replies. */
  readonly items: ComputedRef<readonly ReviewItemView[]>;
  /** The key of the picked thread, or null. */
  readonly activeKey: ComputedRef<string | null>;
  /** Pick a thread, or `null` for none. `{ announce: true }` also asks for the Comments tab. */
  readonly setActive: (key: string | null, options?: ReviewActivationOptions) => boolean;
  /** Add a thread on the current slide, at `target.nodeId`, the selected element, or the slide. */
  readonly comment: (
    text: string,
    author?: string,
    target?: { readonly nodeId: string }
  ) => boolean;
  /** Reply to a thread, or to the thread of a reply. */
  readonly reply: (item: ReviewItemView, text: string, author?: string) => boolean;
  /** Replace the text of a comment or a reply. */
  readonly edit: (item: ReviewItemView, text: string) => boolean;
  /** Delete a reply, or a thread with all of its replies. */
  readonly remove: (item: ReviewItemView) => boolean;
  /** Resolve the thread of `item`. */
  readonly resolve: (item: ReviewItemView) => boolean;
  /** Reopen the thread of `item`. */
  readonly reopen: (item: ReviewItemView) => boolean;
  /** The name that signs a new comment or reply, or null. */
  readonly author: ComputedRef<string | null>;
  /** Set the name that signs later comments and replies (`editor.setAuthor`). */
  readonly setAuthor: (author: string | undefined) => void;
  /** The element that a new comment points at, or null for the slide. */
  readonly commentTarget: ComputedRef<ReviewCommentTarget | null>;
  /** True while no deck is open for editing. */
  readonly readOnly: ComputedRef<boolean>;
  /** False until the editor has a deck. */
  readonly ready: ComputedRef<boolean>;
}

/** Read the comments of the current slide and act on them, in `PptxEditorRoot`. @public */
export function useReview(): UseReviewReturn {
  return useReviewOf(usePptxEditor());
}

/** What `useReviewOf` returns: the same as `useReview`. @public */
export type UseReviewOfReturn = UseReviewReturn;

/** `useReview` over an explicit editor (a ref, a getter, or the instance). @public */
export function useReviewOf(
  editor: MaybeRefOrGetter<PptxEditorInstance | null>
): UseReviewOfReturn {
  const instance = computed(() => toValue(editor));
  const snapshot = shallowRef<EditorSnapshot | null>(null);
  watch(
    instance,
    (e, _old, onCleanup) => {
      snapshot.value = e ? e.snapshot() : null;
      if (!e) return;
      onCleanup(e.subscribe(() => (snapshot.value = e.snapshot())));
    },
    { immediate: true }
  );
  const items = computed(() => reviewItems(instance.value, snapshot.value));
  const actions = (): ReturnType<typeof reviewActions> => reviewActions(instance.value);
  return {
    items,
    activeKey: computed(() => snapshot.value?.activeComment ?? null),
    setActive: (key, options) => actions().setActive(key, options),
    comment: (text, author, target) => actions().comment(text, author, target),
    reply: (item, text, author) => actions().reply(item, text, author),
    edit: (item, text) => actions().edit(item, text),
    remove: (item) => actions().remove(item),
    resolve: (item) => actions().resolve(item),
    reopen: (item) => actions().reopen(item),
    author: computed(() => snapshot.value?.commentAuthor ?? null),
    setAuthor: (author) => actions().setAuthor(author),
    commentTarget: computed(() => {
      const s = snapshot.value;
      return s ? commentTarget(s.selection, s.shapes) : null;
    }),
    readOnly: computed(() => reviewReadOnly(snapshot.value)),
    ready: computed(() => reviewReady(snapshot.value)),
  };
}
