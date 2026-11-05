---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

`PptxEditor` and its root take a `modules` prop for the pro review and collaboration modules, a new `fontFamilies` list updates the font picker in place (`setFontFamilies`), and `save()` returns `Uint8Array<ArrayBuffer>`, so `new Blob([bytes])` type-checks on TypeScript 5.7 and later.
