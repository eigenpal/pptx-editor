---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/editor-api': patch
---

A sync that splits before a read is now one undo step on a live editor: automation batches that carry the same `historyGroup` share one history entry.
