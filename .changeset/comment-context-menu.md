---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
---

Add **New comment** to the canvas context menus, with `registerCommentDraft` and `requestCommentDraft`. Comment commands take `expect` and can fail with `errors.commentChanged`, `errors.commentAmbiguous`, or `errors.noComment`. Threads report `identity`; send it in `expect`. Ids of repeated comments, comments with `~`, and some legacy comments change, so read ids from the current snapshot.
