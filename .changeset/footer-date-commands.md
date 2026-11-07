---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/editor-api': patch
'@pptx-editor.dev/i18n': patch
---

Add `setFooter` and `setDateTime` commands that show or hide the footer and the date on all slides, through the placeholders of each layout or its master. `EditorSnapshot.footer` and `EditorSnapshot.dateTime` report the current slide. A slide that shows a footer or a date already keeps its box and format.
