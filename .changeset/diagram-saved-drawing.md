---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

Draw diagrams (`dgm:` graphic frames) from their saved drawing in the editor and in PDF export. A diagram stays read-only and selects as one object. A drawing past 512 shapes or 8 group levels, or past 4096 shapes and groups on one slide, shows the placeholder. A diagram with no saved drawing still shows a placeholder, and the PDF export still reports `frame.diagram` for it. `ResolvedShape.textXfrm` gives the text rectangle of a diagram shape, and the layout `frame` item gives the laid-out drawing as `drawing`.
