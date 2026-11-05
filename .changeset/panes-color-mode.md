---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

The slide rail and the notes pane take DOCX Editor's pane props (`open`, `defaultOpen`, `onOpenChange`, `className`, and `paneWidth` and `toggle` on the rail), the side panel takes `className` and `paneWidth`, `useNotesPane` and `useEditorColorMode` give the notes state and the resolved theme, and `PptxEditor.Root` takes `colorMode` (`'light'`, `'dark'`, or `'system'`) for the chrome, with `colorMode` and `useEditorColorMode` in Vue too.
