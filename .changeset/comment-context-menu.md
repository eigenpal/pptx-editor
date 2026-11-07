---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
---

Start a comment from the canvas context menu: the text, shape, picture, and table menus show **New comment** (`review.addComment`) while the comments pane of `@pptx-editor.dev/pro` can open a draft. The row opens the **Comments** tab when it is closed and focuses a draft on the right-clicked element. `registerCommentDraft` connects your own comments pane to the row, and `requestCommentDraft` runs it from code. Comment threads report `textEditable`.

Ctrl+Enter or Cmd+Enter in a comment field posts the comment and no longer starts the slide show. A comment that a collaborator adds while another deletes the last thread of the slide is kept. The comments pane stays fast on slides with thousands of comments. Comments that a file stores twice under one id, and legacy comments whose author id holds a hyphen, can be edited and deleted.

If you store comment ids, note these changes. A threaded comment whose GUID the slide repeats, or whose file id holds `~`, now has the id `<slide part>#<GUID>~<mark>~<n>`. The mark comes from the comment's author and creation time. While the GUID repeats, the plain `<slide part>#<GUID>` id names no comment. When a delete leaves one copy of the GUID, the snapshot shows the plain id again, and the marked id still resolves. A legacy comment id encodes `-`, `#`, and `%` in its author and index.

Comment threads and replies report `identity`. Edits, resolves, and deletes of other comments do not change it, so key UI state by it. The comments pane keys its cards, its open fields, and the selected thread by `identity`. When a collaborator edits a comment, its open edit box and an unsent reply keep their text. A field closes without saving only when its comment is deleted.
