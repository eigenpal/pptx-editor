---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

Decks no longer stay on the loading spinner: the open waits at most 2 seconds for fonts. Add `editor.fontMeasurement()`, `openFailureReport()`, and the `showCopy` prop of `PptxEditorOpenFailure`. `DocumentOpenFailureReason` adds `blocked`, so update exhaustive maps of it. Large and damaged decks open faster and more safely.
