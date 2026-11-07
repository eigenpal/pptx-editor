---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/editor-api': patch
'@pptx-editor.dev/i18n': patch
---

Add `setFooter` and `setDateTime` commands that show or hide the footer and the date, through the placeholders of each layout or its master. They and `setSlideNumbers` take `indices` and `hideOnTitleSlides`. Turning one on without `text` keeps what the slides show. `EditorSnapshot.footer` and `EditorSnapshot.dateTime` report what the open slides show. A slide that shows a footer or a date keeps its box and format.
