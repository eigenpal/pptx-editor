---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
'@pptx-editor.dev/i18n': patch
---

Keep the editor responsive on slides with thousands of shapes. On selections of more than 250 shapes, `can()` checks only the editing mode for some commands, and `exec()` makes the full check.
