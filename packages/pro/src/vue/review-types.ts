/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The public prop types of `PptxEditorReview` and its parts: the twin of
// `../react/review-types.ts`. A Vue part takes its content from its default slot, and the pane's
// list render callback is the `#card` scoped slot (`{ item }`).

import type { VNode } from 'vue';
import type { ReviewItemView } from '../review/comment-view.ts';

/** Shared props of every part. @public */
export interface ReviewPartProps {
  /** A class name added to the part's element. */
  className?: string;
  /** Merge the part's wiring onto its single slot element instead of the default element. */
  asChild?: boolean;
  /** Render nothing. Inside the packaged card, this removes the part. */
  hidden?: boolean;
}

/** Props of the action parts, which also take an icon. @public */
export interface ReviewActionProps extends ReviewPartProps {
  /** The icon. Default: the part's Material Symbol; the default slot also replaces it. */
  icon?: VNode | string;
}

/** Props of `PptxEditorReview.Draft`, the compose box of a new comment. @public */
export interface ReviewDraftProps extends ReviewPartProps {
  /** The element that the new comment points at. Default: the selected element, else the slide. */
  nodeId?: string;
}

/** The label resolver of `PptxEditorReview`: a string for a key, or undefined for the default. @public */
export type ReviewTranslate = (key: string) => string | undefined;

/** Props of `PptxEditorReview`. @public */
export interface ReviewProps extends Omit<ReviewPartProps, 'asChild'> {
  /** Label resolver. A key it does not resolve falls back to the bundled catalog. */
  t?: ReviewTranslate;
  /** A class name for each card. */
  card?: { className?: string };
  /** Render the packaged arrangement. `false` renders only the root parts you pass. */
  preset?: boolean;
  /** Show only some threads. Replies follow their thread. */
  filter?: (item: ReviewItemView) => boolean;
}
