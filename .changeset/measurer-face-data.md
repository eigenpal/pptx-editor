---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Strikethrough takes its position and thickness from the font, in the editor and in the PDF export. A double strikethrough centers on the same line. Kerning leaves full-width East Asian text unchanged. Layout measures a character that no font contains at an estimated width, also when the shaper would draw a similar glyph in its place. A missing hyphen or dash takes the width of the font's own hyphen.
