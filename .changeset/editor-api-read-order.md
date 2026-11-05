---
'@pptx-editor.dev/editor-api': patch
---

Every read and lookup now sees the writes queued before it in the same `sync()`, such as `getItemOrNullObject` after an `add` or `getSelectedTextRange()` after `setSelected()`.
