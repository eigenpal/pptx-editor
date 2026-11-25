---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add the `TextFrame.getTextFit()` SDK extension. It reports whether text overflows its shape, from the layout that `getImageAsBase64()` draws, with `overflows`, `textHeight`, `availableHeight`, and `substitutedFonts`. It works on a server runtime with no canvas. The automation protocol adds the `readTextFit` operation.
