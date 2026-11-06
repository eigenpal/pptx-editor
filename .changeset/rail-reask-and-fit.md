---
'@pptx-editor.dev/core': patch
---

A slide rail thumbnail out of view no longer paints when its slide finishes opening or the rail renders again; only a thumbnail asked for the first time paints at once. The canvas fits the slide again when its surface padding (`--doc-slide-padding`) changes. Moving between slides on one layout no longer checks New Slide again, so a slide switch in a 5,000-slide deck costs about a third of before. `registerFontFaces` takes a `family` option that names each face in the font set.
