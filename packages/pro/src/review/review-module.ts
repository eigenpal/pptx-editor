/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Reads PresentationML comments. PresentationML has no tracked-change markup.

import type { EditorModule, ReviewDisplayMode, ReviewItem } from '@pptx-editor.dev/core/editor';
import { rememberLicenseKey, type ProLicenseOptions } from '../license.ts';
import { legacyComments } from './legacy-comments.ts';

/** Options of `reviewModule`. @public */
export interface ReviewModuleOptions extends ProLicenseOptions {
  /** Display modes to unlock. Default: `final` and `markup`. */
  readonly displayModes?: readonly ReviewDisplayMode[];
}

/**
 * The review module. Register it with `createPptxEditor({ modules: [reviewModule()] })`.
 *
 * @public
 */
export function reviewModule(options: ReviewModuleOptions = {}): EditorModule {
  rememberLicenseKey(options.licenseKey);
  return {
    id: 'review',
    review: {
      displayModes: options.displayModes ?? ['final', 'markup'],
      collectReviewItems: ({ package: pkg }): readonly ReviewItem[] => legacyComments(pkg),
    },
  };
}
