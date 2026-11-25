---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `Slide.hidden` to the editing API, an SDK extension. Load it to read whether a slide is hidden, and set it to hide or show the slide. The write changes `p:sld/@show` through the new automation operation `setSlideHidden`, which the `setSlidesHidden` editor command also uses.
