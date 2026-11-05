---
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/core': patch
---

The React and Vue adapters add `useFontFamily()` for a host font picker, export the core types their props and hooks use (such as `PptxEditorInstance`, `EditorCommand`, and `ChromeSlotId`), warn in development when `modules` changes after mount, and no longer export the deprecated `WORD_DEFAULT_FONT`.
