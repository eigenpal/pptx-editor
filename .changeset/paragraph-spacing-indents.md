---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add SDK extensions on `ParagraphFormat` for line spacing, space before and after, and indents. Add the `ParagraphLineSpacingRule` enum and `TableCell.paragraphFormat`. A `150%` spacing value now reads as 1.5 lines. Reads return the value in force, or `null` for a mixed range. The automation `setParagraph` and `readParagraph` operations add the matching fields.
