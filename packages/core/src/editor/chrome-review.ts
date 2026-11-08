// The review controls: New comment (`review.addComment`), after DOCX Editor's
// (docx-editor/packages/core/src/editor/review-chrome-controls.ts,
// packages/core/src/editor/toolbar-commands.ts). The control runs no engine command: a comment
// needs its text first, and the text comes from the draft of a comments pane (`PptxEditorReview` of
// the pro package), the way the link popup owns `text.link`. So the control shows only while a
// comments pane can open a draft (`registerCommentDraft`), and the review module reads comments.
// DOCX Editor commit adf0e178.

import type { Editor } from '../contracts/index.ts';
import type { ChromeGroup, ToolbarCommandState } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { canRequestCommentDraft, requestCommentDraft } from './pane-requests.ts';

/** The review slot ids. @public */
export type ReviewSlotId = 'review.addComment';

/** The `review` group: menus and hosts place its control. @internal */
export const REVIEW_CHROME_GROUP: ChromeGroup = {
  id: 'review',
  labelKey: 'toolbar.groups.review',
  contextual: true,
  controls: [
    {
      id: 'addComment',
      labelKey: 'toolbar.newComment',
      paths: ICON_PATHS['add_comment'] ?? [],
      state: { kind: 'command' },
    },
  ],
};

/**
 * Whether a menu shows the New comment row for `editor`: the review module reads comments, and
 * a comments pane can open a draft (a mounted pane, or a side panel or host that shows the
 * **Comments** tab on request). Without the pro review pane, the row is not in the menu.
 * @internal
 */
export function commentRowShown(editor: Editor | null | undefined): boolean {
  if (!editor) return false;
  const snap = editor.snapshot();
  // Comments are on slides: the slide master view, which shows layouts, has no comment row.
  if (snap.status !== 'ready' || snap.presentationView === 'slideMaster') return false;
  return snap.comments != null && canRequestCommentDraft(editor);
}

/**
 * The state of `review.addComment`: on while the row shows and the deck is open for editing.
 * In view mode the row stays, off with the viewing reason, as DOCX Editor's row does.
 * @internal
 */
export function addCommentState(editor: Editor): ToolbarCommandState {
  const id: ReviewSlotId = 'review.addComment';
  const off = (reason: string): ToolbarCommandState => ({
    id,
    enabled: false,
    disabledReason: reason,
    active: false,
  });
  const snap = editor.snapshot();
  if (snap.status !== 'ready') return off('disabled.noEditor');
  if (!commentRowShown(editor)) return off('disabled.noCommentPane');
  if (snap.mode === 'view') return off('disabled.viewing');
  return { id, enabled: true, disabledReason: null, active: false };
}

/**
 * Run `review.addComment`: ask the comments pane for a new comment draft. The draft points at
 * the selected element, as the pane's own **New comment** does. Nothing about the deck changes
 * until the comment is posted. @internal
 */
export function runAddComment(editor: Editor): { ok: true } | { ok: false; reason: string } {
  const state = addCommentState(editor);
  if (!state.enabled)
    return { ok: false, reason: state.disabledReason ?? 'disabled.noCommentPane' };
  return requestCommentDraft(editor)
    ? { ok: true }
    : { ok: false, reason: 'disabled.noCommentPane' };
}
