---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
'@pptx-editor.dev/pro': minor
---

Add the document refresh API, as in DOCX Editor: `createDocumentRefresh(editor)` shows updated `.pptx` files from your server without the loading screen and flashes what changed. Add the `replaceDocument` command, `getDocumentHandle()`, and `DocumentChange.source` on the `change` event. With `reviewModule()`, results get computed changes; `comparePresentations()` lists them on a server.
