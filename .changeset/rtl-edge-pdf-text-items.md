---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Place right-to-left lines and text columns on whole points, so they sit up to half a point more accurately. In PDF export, draw upright vertical characters upright, and write subscript and superscript runs as their own text items, so text extraction reads them apart.
