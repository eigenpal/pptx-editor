---
'@pptx-editor.dev/core': patch
---

Apply a picked font family to East Asian and complex script text too: the font picker and automation `setFont` write `a:ea` and `a:cs` with `a:latin`.
