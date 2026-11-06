---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
'@pptx-editor.dev/editor-api': patch
---

A long deck with a damaged slide or notes slide is refused at once, wherever the slide is read. After edits, the deck stays open read-only with `snapshot().openFailure`, and commands, undo, save, and automation writes and exports are refused. `snapshot().isOpening` is true until the first slide shows, and `load` accepts any `startSlide` number.
