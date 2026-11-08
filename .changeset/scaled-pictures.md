---
'@pptx-editor.dev/core': patch
---

A slide of many large pictures no longer freezes the editor in browsers that decode pictures on the main thread.

- When the raster pictures of one slide, layout, master, or drawing hold more than about 105 million decoded pixels, each one draws on screen scaled down to its share of that budget. The browser decodes and scales them off the main thread, and the placeholder paints until each one is ready. A slide within the budget draws its pictures as stored. After an edit, scaled sizes stay until a picture's share moves by more than 25%, and a picture keeps its last scaled size until the new size is ready. Export, copy, and save use the stored pictures.
- Before, 200 pictures of 16 million pixels each on one slide stopped the page for up to a minute, because such a browser decodes each picture at its stored size while it paints it.
