---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pptx-to-pdf': patch
---

`runGlyphs` gives each character of a glyph run with the left edge of its cell, in source order, and draws each mirrored character (UAX #9 L4, such as a bracket) as its mirror in a right-to-left run. Warped text, the PDF export, and the empty-item bullet hint place right-to-left text in reading order with it.
