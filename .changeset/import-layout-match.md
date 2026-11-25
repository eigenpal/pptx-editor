---
'@pptx-editor.dev/core': patch
---

`insertSlidesFromBase64` with `UseDestinationTheme` picks a better layout when several layouts share a type. It prefers the same layout name, then the same master name, then the same layout when both names match. Thus a slide from `Slide.exportAsBase64` keeps its layout in the same presentation, when its layout and master have names.
