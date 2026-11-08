// The deck's review authors and the color each one draws in, as DOCX Editor's
// `useReviewAuthors` (docx-editor/packages/react/src/editor/useReviewAuthors.ts at f189d6503).
// The Vue twin is `useReviewAuthors` in `../../vue/src/composables.ts`.

import { useCallback } from 'react';
import type { ReviewAuthorInfo } from '@pptx-editor.dev/core/editor';
import { useEditorState, usePptxEditor } from './context.ts';

const NONE: readonly ReviewAuthorInfo[] = Object.freeze([]);

/**
 * The review authors of the open deck, in slot order (`getReviewAuthors`): every comment author
 * by first appearance, then the author of new comments, each with the color the comment markers,
 * the text bands, and the comments pane draw them in. Renders again only when a slot changes.
 * Use it for a legend or for your own comment UI.
 *
 * ```tsx
 * const authors = useReviewAuthors();
 * authors.map(({ author, color }) => (
 *   <li key={author} style={{ borderColor: color }}>{author}</li>
 * ));
 * ```
 *
 * @public
 */
export function useReviewAuthors(): readonly ReviewAuthorInfo[] {
  const editor = usePptxEditor();
  return useEditorState(useCallback(() => editor?.getReviewAuthors() ?? NONE, [editor]));
}
