---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Fix tiled picture fills in printed slides. The editor draws each tile as an image, so a printed slide keeps the picture's pixels and the tile positions, and mirrored tiles show no seams. Tiles fill the whole shape, also a callout pointer outside the shape's box. A PNG or BMP resolution stated in pixels per metre that is a rounded whole DPI reads as that DPI, so a 144 DPI tile is exactly 0.5 in wide in the editor and in PDF export.
