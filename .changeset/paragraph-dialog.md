---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

Add the Paragraph dialog: `setParagraphFormat` sets alignment, indents, space before and after, and line spacing as one undo step, `PptxEditor.ParagraphDialog` (with the dialog named parts and `useParagraphDialog`) opens from Format > Paragraph…, Line spacing options…, and the text context menu, `useParagraphFormat` reads and writes the same fields in React and Vue, with its `ParagraphFormatRead` and `ParagraphFormatCommand` types, and `setLineSpacing` now refuses a multiple below 0.01 or above 9.99, as the dialog does.
