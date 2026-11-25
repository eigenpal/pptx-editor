---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/editor-api': minor
---

Add `Presentation.replaceAll(searchText, replacementText, options)` to the editing API, with `SearchOptions` (`matchCase`, `matchWholeWord`). It replaces text in slide shapes, table cells, and speaker notes, keeps the font of the replaced text, and returns the count. The automation host adds the `replaceText` operation. Find no longer matches a field through U+FFFC.
