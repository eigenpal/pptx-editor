/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The review module: the comment threads of the deck, in both comment formats, as the review
// queue, the snapshot's comment threads, and the slide's comment markers. The comments pane is
// `PptxEditorReview` in the `/react` and `/vue` entries. PresentationML has no tracked-change
// markup.

import {
  readCommentThreads,
  type EditorModule,
  type ReviewDisplayMode,
  type ReviewItem,
  type ReviewPaneOptions,
} from '@pptx-editor.dev/core/editor';
import { presentationInfo, type PptxPackage } from '@pptx-editor.dev/core/store';
import { rememberLicenseKey, type ProLicenseOptions } from '../license.ts';

/** Options of `reviewModule`. @public */
export interface ReviewModuleOptions extends ProLicenseOptions {
  /** Display modes to unlock. Default: `final` and `markup`. */
  readonly displayModes?: readonly ReviewDisplayMode[];
  /**
   * The comment view settings, as DOCX Editor's `pane` option: `commentMarkers` (`'icon'`, the
   * default, or `'initials'`) sets how a comment marker on the slide looks, and `opening`
   * (`'auto'`, the default, or `'manual'`) sets whether a press on a marker asks for the
   * **Comments** tab of the side panel. An unknown field or value throws a `TypeError`.
   */
  readonly pane?: ReviewPaneOptions;
}

const PANE_VALUES: { readonly [K in keyof ReviewPaneOptions]-?: readonly string[] } = {
  opening: ['auto', 'manual'],
  commentMarkers: ['initials', 'icon'],
};

/** Check `pane` where the host wrote it, so a misspelled setting fails there. */
function checkPane(pane: ReviewPaneOptions | undefined): ReviewPaneOptions | undefined {
  if (pane === undefined) return undefined;
  for (const [key, value] of Object.entries(pane)) {
    const allowed = PANE_VALUES[key as keyof ReviewPaneOptions];
    if (!allowed) throw new TypeError(`reviewModule: unknown pane option "${key}"`);
    if (value !== undefined && !allowed.includes(value as string))
      throw new TypeError(
        `reviewModule: pane.${key} must be one of ${allowed.map((v) => `"${v}"`).join(', ')}`
      );
  }
  return Object.freeze({ ...pane });
}

/** Every comment and reply of the deck, in slide order and then thread order. */
function reviewItems(pkg: PptxPackage): ReviewItem[] {
  const out: ReviewItem[] = [];
  for (const slide of presentationInfo(pkg).slides)
    for (const thread of readCommentThreads(pkg, slide.partName)) {
      const item = (c: typeof thread | (typeof thread.replies)[number], parentId?: string) =>
        out.push({
          id: c.id,
          kind: 'comment',
          slidePart: slide.partName,
          text: c.text,
          ...(c.author ? { author: c.author } : {}),
          ...(c.date !== undefined ? { date: c.date } : {}),
          ...(parentId ? { parentId } : {}),
        });
      item(thread);
      for (const reply of thread.replies) item(reply, thread.id);
    }
  return out;
}

/**
 * The review module. Register it with `createPptxEditor({ modules: [reviewModule()] })`: the
 * editor snapshot then reports the current slide's comment threads (`comments`), the slide shows
 * a marker for each thread, and `collectReviewItems` reads every comment of the deck. Render
 * `PptxEditorReview` from `@pptx-editor.dev/pro/react` or `/vue` for the comments pane.
 *
 * @public
 */
export function reviewModule(options: ReviewModuleOptions = {}): EditorModule {
  const pane = checkPane(options.pane);
  rememberLicenseKey(options.licenseKey);
  return {
    id: 'review',
    review: {
      displayModes: options.displayModes ?? ['final', 'markup'],
      collectReviewItems: ({ package: pkg }): readonly ReviewItem[] => reviewItems(pkg),
      ...(pane ? { pane } : {}),
    },
  };
}

export type {
  CommentMarkerStyle,
  ReviewPaneOpening,
  ReviewPaneOptions,
} from '@pptx-editor.dev/core/editor';
