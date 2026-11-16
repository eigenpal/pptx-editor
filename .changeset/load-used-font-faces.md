---
'@pptx-editor.dev/core': minor
---

A deck shows its first slide sooner because the editor first loads only the font faces that the slide's text uses, and loads the other faces in the background. Font resolvers receive these faces in the new `faces` field of the request, so `packagedFonts()` and `googleFonts()` from `@docx-editor.dev/fonts` 2.38 load the same way.
