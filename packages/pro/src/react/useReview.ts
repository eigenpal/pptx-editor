/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The comments of the current slide, as a hook, after DOCX Editor's `useReview`
// (docx-editor/packages/pro/src/react/useReview.ts at adf0e178c). THIS is the API:
// `PptxEditorReview` and its parts are one rendering of what it returns, and nothing is reachable
// through them that is not reachable here. The Vue twin is `../vue/useReview.ts`.

import { useCallback, useMemo, useSyncExternalStore } from 'react';
import type {
  CommentTextRange,
  EditorSnapshot,
  PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { usePptxEditor } from '@pptx-editor.dev/react';
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
 * What `useReview` returns: the comments of the current slide and the things a card can do.
 * Every action reports whether the editor applied it. A refused action changes nothing, and the
 * editor publishes its reason (`lastRejection`), which the edit notice shows.
 *
 * @public
 */
export interface UseReviewReturn {
  /**
   * The comments of the current slide, in file order: each thread, then its replies. Empty
   * without a review module or a deck.
   */
  readonly items: readonly ReviewItemView[];
  /** The key of the picked thread (a marker press, `setActive`), or null. */
  readonly activeKey: string | null;
  /**
   * Pick a thread, as a press on its marker does, or `null` for none. With `{ announce: true }`
   * it also asks for the **Comments** tab of the side panel. False for a key that is not a thread
   * of the current slide.
   */
  readonly setActive: (key: string | null, options?: ReviewActivationOptions) => boolean;
  /**
   * Add a thread on the current slide. It points at `target.range` of the text of
   * `target.nodeId`, at `target.nodeId`, else at the selected text or element
   * (`commentTarget`), else at the slide. `author` overrides `author` for this comment.
   */
  readonly comment: (
    text: string,
    author?: string,
    target?: { readonly nodeId: string; readonly range?: CommentTextRange }
  ) => boolean;
  /** Reply to a thread, or to the thread of a reply. */
  readonly reply: (item: ReviewItemView, text: string, author?: string) => boolean;
  /** Replace the text of a comment or a reply. False when `item.editable` is false. */
  readonly edit: (item: ReviewItemView, text: string) => boolean;
  /** Delete a reply, or a thread with all of its replies. */
  readonly remove: (item: ReviewItemView) => boolean;
  /** Resolve the thread of `item`. A resolved thread reports true with no write. */
  readonly resolve: (item: ReviewItemView) => boolean;
  /** Reopen the thread of `item`. An open thread reports true with no write. */
  readonly reopen: (item: ReviewItemView) => boolean;
  /** The name that signs a new comment or reply, or null. Without one, `comment` is refused. */
  readonly author: string | null;
  /** Set the name that signs later comments and replies (`editor.setAuthor`). */
  readonly setAuthor: (author: string | undefined) => void;
  /** The element that a new comment points at, or null for the slide. */
  readonly commentTarget: ReviewCommentTarget | null;
  /** True while no deck is open for editing: in view mode, and before a deck opens. */
  readonly readOnly: boolean;
  /** False until the editor has a deck. */
  readonly ready: boolean;
}

/**
 * Read the comments of the current slide and act on them, in the editor root. For headless or
 * server work, use the comment commands of `@pptx-editor.dev/core` or `@pptx-editor.dev/editor-api`.
 *
 * @public
 */
export function useReview(): UseReviewReturn {
  return useReviewOf(usePptxEditor());
}

/** `useReview` over an explicit editor, for a host that holds its own. @public */
export function useReviewOf(editor: PptxEditorInstance | null): UseReviewReturn {
  const subscribe = useCallback(
    (notify: () => void) => (editor ? editor.subscribe(notify) : () => undefined),
    [editor]
  );
  const snapshot: EditorSnapshot | null = useSyncExternalStore(
    subscribe,
    () => (editor ? editor.snapshot() : null),
    () => null
  );
  const items = useMemo(() => reviewItems(editor, snapshot), [editor, snapshot]);
  const actions = useMemo(() => reviewActions(editor), [editor]);
  const selection = snapshot?.selection;
  const shapes = snapshot?.shapes;
  const target = useMemo(
    () => (selection ? commentTarget(selection, shapes) : null),
    [selection, shapes]
  );
  const activeKey = snapshot?.activeComment ?? null;
  const author = snapshot?.commentAuthor ?? null;
  const readOnly = reviewReadOnly(snapshot);
  const ready = reviewReady(snapshot);
  return useMemo(
    () => ({
      items,
      activeKey,
      ...actions,
      author,
      commentTarget: target,
      readOnly,
      ready,
    }),
    [items, activeKey, actions, author, target, readOnly, ready]
  );
}
