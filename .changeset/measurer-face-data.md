---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Strikethrough takes its position and thickness from the font, in the editor and in the PDF export. A double strikethrough centers on the same line. Kerning leaves full-width East Asian text unchanged. A character that a font does not contain takes a fallback font, also when the shaper would draw a similar glyph in its place.
