/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The public prop types of `PptxEditorReview` and its parts, adapted from DOCX Editor's
// (docx-editor/packages/pro/src/react/review-types.ts at adf0e178c). The Vue twin is
// `../vue/review-types.ts`.

import type { ReactNode } from 'react';
import type { ReviewItemView } from '../review/comment-view.ts';

/** Shared props of every part. @public */
export interface ReviewPartProps {
  /** A class name added to the part's element. */
  className?: string;
  /** Merge the part's wiring onto its single child element instead of the default element. */
  asChild?: boolean;
  /** Render nothing. Inside the packaged card, this removes the part. */
  hidden?: boolean;
  children?: ReactNode;
}

/** Props of the action parts, which also take an icon. @public */
export interface ReviewActionProps extends ReviewPartProps {
  /** The icon. Default: the part's Material Symbol; `children` also replace it. */
  icon?: ReactNode;
}

/** Props of `PptxEditorReview.Draft`, the compose box of a new comment. @public */
export interface ReviewDraftProps extends ReviewPartProps {
  /** The element that the new comment points at. Default: the selected element, else the slide. */
  nodeId?: string;
}

/** The label resolver of `PptxEditorReview`: a string for a key, or undefined for the default. @public */
export type ReviewTranslate = (key: string) => string | undefined;

/** Props of `PptxEditorReview`. @public */
export interface ReviewProps extends Omit<ReviewPartProps, 'children' | 'asChild'> {
  /**
   * Label resolver, as the toolbar and menus take one. A key it does not resolve falls back to
   * the bundled catalog. Placeholders such as `{author}` are filled in.
   */
  t?: ReviewTranslate;
  /** A class name for each card. `className` styles the pane; this styles the cards in it. */
  card?: { className?: string };
  /**
   * The parts of the pane and of its cards. Root parts (`List`, `AddComment`, `Draft`) put as
   * direct children replace the packaged ones. Other children become the card template of the
   * list. A function is the list's render callback, with one call for each thread.
   */
  children?: ReactNode | ((item: ReviewItemView) => ReactNode);
  /**
   * Render the packaged arrangement: the **New comment** button, the compose box, and the list.
   * `false` renders only the root parts you pass, with the pane's data and context.
   */
  preset?: boolean;
  /** Show only some threads, for example the open ones. Replies follow their thread. */
  filter?: (item: ReviewItemView) => boolean;
}
