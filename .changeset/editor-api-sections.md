---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add the SDK extension `Presentation.sections`. Read each section's id, name, position, and slide ids. Add a section at a slide, rename it, move it with its slides, and delete it with or without its slides. Each sync is one undo step. The core automation protocol adds section operations and the `section` handle kind.
