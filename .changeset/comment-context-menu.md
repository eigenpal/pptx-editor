---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
---

Start a comment from the canvas context menu: the text, shape, picture, and table menus show **New comment** (`review.addComment`) while the comments pane of `@pptx-editor.dev/pro` can open a draft. The row opens the **Comments** tab when it is closed and focuses a draft on the right-clicked element. `registerCommentDraft` connects your own comments pane to the row, and `requestCommentDraft` runs it from code. Comment threads report `textEditable`.

Ctrl+Enter or Cmd+Enter in a comment field posts the comment and no longer starts the slide show. A comment that a collaborator adds while another deletes the last thread of the slide is kept. The comments pane stays fast on slides with thousands of comments. Comments that a file stores twice under one id, and legacy comments whose author id holds a hyphen, can be edited and deleted.

If you store comment ids, note these changes. A threaded comment whose GUID the slide repeats, or whose file id holds `~`, now has the id `<slide part>#<GUID>~<mark>`. The mark comes from the comment's author and creation time. Twins, comments that also share the mark, add their place: `<slide part>#<GUID>~<mark>~<n>`. While the GUID repeats, the plain `<slide part>#<GUID>` id names no comment, and while twins share a mark, the id without a place names none. When a delete leaves one copy, the snapshot shows the shorter id, and the longer id still resolves. A legacy comment id encodes `-`, `#`, and `%` in its author and index. A legacy comment whose author and index the slide repeats has the id `<slide part>#cm-<author>-<idx>-<n>`, and its plain id names no comment while they repeat.

The commands that act on a comment (`replyToComment`, `editComment`, `deleteComment`, `resolveComment`, and `reopenComment`) take `expect`: the author, date, and text that you showed. The editor refuses the command with `errors.commentChanged` when the comment at `commentId` has another author or date. For an id that holds a place among twins (comments that share a GUID, author, and creation time, or legacy comments that share an author and index), it also compares the text, and it refuses the command when no twin is left. A collaborator's edit of a comment without a twin does not refuse your edit.

Comment threads and replies report `identity`, a token for UI state. Once shown, an identity never names another comment. A comment that never had a twin keeps it until it is deleted. Twins keep it when one change explains their group: an edit that changes at most one text, or one twin more or less at one place. Otherwise every twin of the group takes a new identity.

The comments pane keys its cards, its open fields, and the selected thread by `identity`, and sends `expect` with each action. When a collaborator edits a comment, its open edit box and an unsent reply keep their text. When the editor refuses a save or a reply, the field keeps its text and says why. After a notice that the comment changed, a second save or reply applies to the comment as it shows now. A field closes without saving when its comment is deleted, or when its twins take new identities.
