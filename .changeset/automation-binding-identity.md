---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/editor-api': patch
---

A binding whose shape or slide is deleted no longer binds a new shape that takes the same id; an undo that restores the shape makes it work again.
