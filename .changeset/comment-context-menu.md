---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
---

Add **New comment** to the canvas context menus, with `registerCommentDraft` and `requestCommentDraft`. Comment commands take `expect` and can fail with `errors.commentChanged`, `errors.commentAmbiguous`, or `errors.noComment`. Threads report `identity`. Ids of repeated comments, comments with `~`, and legacy comments with `-`, `#`, or `%` change, so read ids from the current snapshot.
