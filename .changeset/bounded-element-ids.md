---
'@pptx-editor.dev/core': patch
---

New shapes, slides, and copied comments get ids inside their schema range and unique on the slide, also when the file holds ids such as `4294967295` or `1e400`.
