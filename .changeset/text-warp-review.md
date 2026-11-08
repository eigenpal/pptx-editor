---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Warped text stops early and draws flat when its glyphs need too many path commands. Pour warps with a zero inner radius and can warps with a zero arc height now draw warped. Removing an inherited warp writes `textNoShape`. Picture bullets draw on warped text, and the PDF export strokes the text outline of warped glyphs.
