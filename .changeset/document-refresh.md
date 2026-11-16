---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
'@pptx-editor.dev/pro': minor
---

Add the document refresh API, as in DOCX Editor. `createDocumentRefresh(editor)` shows updated `.pptx` files from your server without a reset of the view: the loading screen does not show, and only the slides and thumbnails that changed paint again. `highlightChanges()` flashes the thumbnails of changed slides and the changed shapes and text on the slide on screen. Add `replaceDocument`, which replaces the presentation with a file as one undoable edit, and `getDocumentHandle()`. The `change` event receives a `DocumentChange` whose `source` names a load or a document refresh. With `reviewModule()`, a result without `changes` gets computed changes: slides, shapes, text, notes, layouts, masters, and themes. `comparePresentations()` lists the same changes on a server.
