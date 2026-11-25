---
'@pptx-editor.dev/editor-api': patch
---

`editor-api` declares `@napi-rs/canvas` `^1.0.12` as an optional peer dependency. Earlier releases draw the SVG of a `getImageAsBase64()` picture without its `<image>` elements. The slide pictures are then missing with no error. The runtime guide states this minimum. If your package manager reports a peer conflict, update `@napi-rs/canvas`.
