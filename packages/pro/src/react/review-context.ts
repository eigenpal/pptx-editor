/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The pane's private contexts and the two public hooks that read them, after DOCX Editor's
// (docx-editor/packages/pro/src/react/review-context.ts at adf0e178c). The Vue twin is
// `../vue/review-context.ts`.

import { createContext, useCallback, useContext } from 'react';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { useTranslation } from '@pptx-editor.dev/react';
import {
  reviewAuthorInfo,
  reviewActions,
  type ReviewAuthorInfo,
  type ReviewCommentTarget,
  type ReviewItemView,
} from '../review/comment-view.ts';
import type { UseReviewReturn } from './useReview.ts';
import type { ReviewTranslate } from './review-types.ts';

/** What the pane gives its parts. */
export interface ReviewPaneValue {
  /** The host's label resolver, when it passed one. Parts read it through `useReviewLabel`. */
  readonly t: ReviewTranslate | undefined;
  /** A card's class name, from the pane's `card` prop. */
  readonly cardClassName: string | undefined;
  /** The review, with `items` narrowed by the pane's `filter`. */
  readonly review: UseReviewReturn;
  /** Items by id, so a card renders its replies without a walk of the list. */
  readonly byId: ReadonlyMap<string, ReviewItemView>;
  /** Whether the compose box of a new comment is open. */
  readonly drafting: boolean;
  /**
   * The element a draft from the canvas menus' New comment row points at, or null: the draft then
   * points at the selected element.
   */
  readonly draftTarget: ReviewCommentTarget | null;
  /** Open the compose box. */
  readonly beginDraft: () => void;
  /** Close it, posted or not. */
  readonly endDraft: () => void;
  /** The `identity` of the comment whose text is open for an edit, or null. */
  readonly editing: string | null;
  /** Open the edit of a comment, or close it with null. */
  readonly setEditing: (key: string | null) => void;
  /** Move the focus to a card after its form closes, so the keyboard user keeps their place. */
  readonly focusCard: (key: string) => void;
}

/** The pane's data, provided once by the root. */
export const ReviewContext = createContext<ReviewPaneValue | null>(null);
/** The item that the surrounding card renders. */
export const ReviewItemContext = createContext<ReviewItemView | null>(null);

/**
 * The comment that the surrounding card (or reply) renders, or null outside one. A host's own
 * card content reads its item here, as the packaged parts do.
 *
 * @public
 */
export function useReviewItem(): ReviewItemView | null {
  return useContext(ReviewItemContext);
}

/**
 * The look of one comment author: the color slot and the CSS color that the canvas marker and
 * the packaged cards use, or `undefined` without an author. A custom card draws with it, so it
 * keeps the colors of the rest of the review chrome.
 *
 * @public
 */
export function useReviewAuthor(author: string | undefined): ReviewAuthorInfo | undefined {
  return author ? reviewAuthorInfo(author) : undefined;
}

const INERT_REVIEW: UseReviewReturn = {
  items: [],
  activeKey: null,
  ...reviewActions(null),
  author: null,
  commentTarget: null,
  readOnly: true,
  ready: false,
};

const INERT_PANE: ReviewPaneValue = {
  t: undefined,
  cardClassName: undefined,
  review: INERT_REVIEW,
  byId: new Map(),
  drafting: false,
  draftTarget: null,
  beginDraft: () => {},
  endDraft: () => {},
  editing: null,
  setEditing: () => {},
  focusCard: () => {},
};

/** The pane's data. A part rendered outside the pane by mistake renders nothing, not an error. */
export function usePane(): ReviewPaneValue {
  return useContext(ReviewContext) ?? INERT_PANE;
}

/** Placeholder values for a label, such as `{ author: 'Ada' }` for `{author}`. */
export type ReviewLabelParams = Readonly<Record<string, string | number>>;

/** A part's strings: the host's `t`, else the bundled catalog. */
export function useReviewLabel(): (key: TranslationKey, params?: ReviewLabelParams) => string {
  const host = useContext(ReviewContext)?.t;
  const { t } = useTranslation();
  return useCallback(
    (key: TranslationKey, params?: ReviewLabelParams) => {
      const own = host?.(key);
      if (own === undefined) return t(key, params);
      return params
        ? own.replace(/\{(\w+)\}/g, (whole, name: string) =>
            Object.hasOwn(params, name) ? String(params[name]) : whole
          )
        : own;
    },
    [host, t]
  );
}
