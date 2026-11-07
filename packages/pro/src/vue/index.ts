/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `@pptx-editor.dev/pro/vue`: the comments pane and the review composables for the Vue adapter,
 * and the review module, so a host imports one path, as DOCX Editor's `@docx-editor.dev/pro/vue`.
 * `@pptx-editor.dev/core` and `@pptx-editor.dev/vue` stay peers, so one engine instance loads.
 *
 * @packageDocumentation
 */

export {
  reviewModule,
  type CommentMarkerStyle,
  type ReviewModuleOptions,
  type ReviewPaneOpening,
  type ReviewPaneOptions,
} from '../review/review-module.ts';
export { type ProLicenseOptions } from '../license.ts';
export { PptxEditorReview, type PptxEditorReviewNamespace } from './PptxEditorReview.tsx';
export type {
  ReviewActionProps,
  ReviewDraftProps,
  ReviewPartProps,
  ReviewProps,
  ReviewTranslate,
} from './review-types.ts';
export {
  useReviewAuthor,
  useReviewItem,
  type UseReviewAuthorReturn,
  type UseReviewItemReturn,
} from './review-context.ts';
export {
  useReview,
  useReviewOf,
  type UseReviewOfReturn,
  type UseReviewReturn,
} from './useReview.ts';
export type {
  ReviewActivationOptions,
  ReviewAuthorInfo,
  ReviewCommentTarget,
  ReviewItemView,
} from '../review/comment-view.ts';
