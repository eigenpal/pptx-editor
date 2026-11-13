---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Fix tiled picture fills in printed slides: each tile prints as an image with no seams between mirrored tiles, and tiles fill the whole shape, up to 4,096 tile images per slide. A PNG or BMP resolution in pixels per meter that rounds to a whole DPI reads as that DPI.
