---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/editor-api': minor
---

Add `Presentation.replaceAll(searchText, replacementText, options)` and `Slide.replaceAll` to the editing API, with `SearchOptions` (`matchCase`, `matchWholeWord`). It replaces text in shapes, table cells, and notes. It keeps the font and returns the count. The automation host adds the `replaceText` operation. Find no longer matches a field through U+FFFC.
