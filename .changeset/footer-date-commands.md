---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/editor-api': patch
'@pptx-editor.dev/i18n': patch
---

Add the `setFooter` and `setDateTime` commands, and `EditorSnapshot.footer` and `EditorSnapshot.dateTime`. `setSlideNumbers` takes `indices` and `hideOnTitleSlides`, and refuses `show: false` when no slide can have a slide number.
