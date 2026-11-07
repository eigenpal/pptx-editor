---
'@pptx-editor.dev/core': patch
---

Paint oversized grapheme clusters, including bullets and empty-list caret markers, in bounded pieces to prevent native text rendering from blocking input. Preserve every source character, character position, and semantic caret boundary. Report the limited shaping across paint pieces as an approximation.
