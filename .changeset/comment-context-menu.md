---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/i18n': minor
---

Add **New comment** to the canvas context menus, with `registerCommentDraft` and `requestCommentDraft`. Comment commands take `expect`, and threads report `identity`. Comments with repeated ids get new ids, so read ids from the current snapshot.
