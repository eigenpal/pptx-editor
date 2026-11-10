---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

Accept font configurations, promises, and ordered resolvers in PDF export. Replace `googleFonts: true` with `fonts: googleFonts({ cjk: true })`. Raw font arrays now require `createFontSource` and `{ sources }`. Export waits for deferred fonts and shares the editor's font validation.
