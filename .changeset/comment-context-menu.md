---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
---

Start a comment from the canvas context menu: the text, shape, picture, and table menus show **New comment** (`review.addComment`) while the comments pane of `@pptx-editor.dev/pro` can open a draft. The row opens the **Comments** tab when it is closed and focuses a draft on the right-clicked element. `registerCommentDraft` connects your own comments pane to the row, and `requestCommentDraft` runs it from code. Comment threads report `textEditable`.

Ctrl+Enter or Cmd+Enter in a comment field posts the comment and no longer starts the slide show. A comment that a collaborator adds while another deletes the last thread of the slide is kept. The comments pane stays fast on slides with thousands of comments. Comments that a file stores twice under one id, and legacy comments whose author id holds a hyphen, can be edited and deleted.
