---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add the SDK extension `Slide.speakerNotes`. Load it to read the notes as plain text, and set it to replace them, one paragraph for each line. A write to a slide with no notes creates the notes slide from the notes master. The core automation protocol adds the `readSlideNotes` and `setSlideNotes` operations.
