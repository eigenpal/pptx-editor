---
'@pptx-editor.dev/pro': patch
'@pptx-editor.dev/core': patch
---

Two peers who type at once into a table cell with no text body both keep their text, a blank second text body in a cell no longer adds an empty line, and the automation API reads and edits the text of every text body of such a cell.
