---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

End right-to-left lines on the whole-point grid of the line model. In PDF export, draw upright vertical characters upright, and write subscript and superscript runs as their own text items, so text extraction reads them apart from the text around them.
