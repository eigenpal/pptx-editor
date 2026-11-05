---
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': patch
---

Both adapters export the same props and return types, such as `PptxEditorViewportProps`, `FindBarProps`, `LoadingSpinnerProps`, `UseEditorCommandReturn`, and `UseSlideshowReturn` (the React `*Handle` names stay as deprecated aliases), and `PptxEditor` takes `locale`, `fontFamilies`, and `onChange` as `PptxEditor.Root` does.
