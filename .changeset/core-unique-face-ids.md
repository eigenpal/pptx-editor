---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/editor-api': patch
---

Font face ids are now unique for each open presentation, so a rasterizer's font registry never mixes the faces of two presentations.
