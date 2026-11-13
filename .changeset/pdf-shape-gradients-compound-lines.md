---
'@pptx-editor.dev/pptx-to-pdf': patch
---

The PDF export draws `path="shape"` gradients as bands of the outline, applies gradient stop alpha once to rectangle and shape path gradients, draws the bands of dashed compound lines, and fades soft edges. The export and the editor share these rules.
