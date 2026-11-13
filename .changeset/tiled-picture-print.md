---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Fix tiled picture fills in printed slides and PDF export. The editor draws each tile as an image, so a printed slide keeps the picture's pixels and the tile positions. Mirrored tiles no longer show seams. A PNG or BMP resolution stated in pixels per metre reads as whole DPI, so a 144 DPI tile is exactly 0.5 in wide in the editor and in PDF export.
