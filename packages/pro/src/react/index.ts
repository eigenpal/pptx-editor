/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `@pptx-editor.dev/pro/react`: the comments pane and the review hooks for the React adapter, and
 * the review module, so a host imports one path, as DOCX Editor's `@docx-editor.dev/pro/react`.
 * `@pptx-editor.dev/core` and `@pptx-editor.dev/react` stay peers, so one engine instance loads.
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
export {
  PptxEditorReview,
  type PptxEditorReviewNamespace,
  type ReviewListProps,
} from './PptxEditorReview.tsx';
export type {
  ReviewActionProps,
  ReviewDraftProps,
  ReviewPartProps,
  ReviewProps,
  ReviewTranslate,
} from './review-types.ts';
export { useReviewAuthor, useReviewItem } from './review-context.ts';
export { useReview, useReviewOf, type UseReviewReturn } from './useReview.ts';
export type {
  ReviewActivationOptions,
  ReviewAuthorInfo,
  ReviewCommentTarget,
  ReviewItemView,
} from '../review/comment-view.ts';
