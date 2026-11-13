---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Fix the right edge of right-to-left paragraphs: their lines end at the column width in whole points. In PDF export, draw upright vertical characters upright, and write subscript and superscript runs as their own text items, so text extraction reads them apart from the text around them.
