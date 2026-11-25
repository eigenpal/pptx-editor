---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add the SDK extension `Slide.duplicate()`. It copies a slide directly after its source in the same presentation and returns the new `Slide`, which works in the same batch. The copy keeps the layout, master, and pictures, and gets its own notes, charts, and tags. The core automation protocol adds the `duplicateSlide` operation.
