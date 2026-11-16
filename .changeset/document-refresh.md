---
'@pptx-editor.dev/core': minor
---

Add `replaceDocument`, which replaces the presentation with a `.pptx` file as one undoable edit, and `getDocumentHandle()`. The `change` event receives a `DocumentChange` whose `source` names a load or a document refresh.
